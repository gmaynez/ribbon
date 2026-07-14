using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    public sealed class RibbonSidebarControl : UserControl
    {
        private readonly VstoHostRuntime _runtime;
        private readonly RibbonPalette _palette;
        private readonly RibbonComboBox _agents;
        private readonly RibbonButton _manage;
        private readonly RichTextBox _transcript;
        private readonly TextBox _prompt;
        private readonly RibbonButton _send;
        private readonly RibbonButton _cancel;
        private readonly Label _status;
        private readonly RibbonStatusDot _statusDot;
        private Font _transcriptRegularFont;
        private Font _transcriptBoldFont;
        private string _sessionId;
        private bool _loaded;
        private bool _hasConversation;

        public RibbonSidebarControl(VstoHostRuntime runtime)
        {
            _runtime = runtime ?? throw new ArgumentNullException(nameof(runtime));
            _palette = RibbonPalette.Detect();
            _agents = new RibbonComboBox(_palette);
            _manage = new RibbonButton(_palette, RibbonButtonKind.Secondary) { Text = "Agents", Glyph = RibbonGlyph.Agents, Width = 88 };
            _transcript = new RichTextBox();
            _prompt = new TextBox();
            _send = new RibbonButton(_palette, RibbonButtonKind.Primary) { Text = "Send", Glyph = RibbonGlyph.Send, Width = 80 };
            _cancel = new RibbonButton(_palette, RibbonButtonKind.Ghost) { Text = "Stop", Glyph = RibbonGlyph.Stop, Width = 76 };
            _status = new Label();
            _statusDot = new RibbonStatusDot { DotColor = _palette.MutedText };

            Dock = DockStyle.Fill;
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = _palette.Background;
            ForeColor = _palette.Text;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
            MinimumSize = new Size(300, 360);
            BuildLayout();
            ShowWelcomeMessage();
            _runtime.SessionUpdate += RuntimeOnSessionUpdate;
        }

        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (_loaded) return;
            _loaded = true;
            await InitializeAsync();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _runtime.SessionUpdate -= RuntimeOnSessionUpdate;
                _transcriptBoldFont?.Dispose();
                _transcriptRegularFont?.Dispose();
            }
            base.Dispose(disposing);
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(12),
                BackColor = _palette.Background
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 124));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));

            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildTranscript(), 0, 1);
            root.Controls.Add(BuildComposer(), 0, 2);
            root.Controls.Add(BuildFooter(), 0, 3);
            Controls.Add(root);
        }

        private Control BuildHeader()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10), Padding = new Padding(12, 10, 12, 10) };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = _palette.Surface };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var brandRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, BackColor = _palette.Surface };
            brandRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40));
            brandRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            var mark = new RibbonBrandMark(_palette) { Margin = new Padding(0, 1, 8, 0) };
            var titles = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = _palette.Surface };
            titles.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            titles.RowStyles.Add(new RowStyle(SizeType.Absolute, 18));
            var title = LabelFor("Ribbon", 10.5f, FontStyle.Bold, _palette.Text);
            var subtitle = LabelFor((_runtime.Registration.HostKind ?? "Office") + " agent workspace", 8f, FontStyle.Regular, _palette.MutedText);
            titles.Controls.Add(title, 0, 0);
            titles.Controls.Add(subtitle, 0, 1);
            brandRow.Controls.Add(mark, 0, 0);
            brandRow.Controls.Add(titles, 1, 0);

            var agentRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, BackColor = _palette.Surface };
            agentRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            agentRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 94));
            agentRow.RowStyles.Add(new RowStyle(SizeType.Absolute, 18));
            agentRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var agentLabel = LabelFor("ACTIVE AGENT", 7.5f, FontStyle.Bold, _palette.MutedText);
            agentRow.Controls.Add(agentLabel, 0, 0);
            agentRow.SetColumnSpan(agentLabel, 2);
            _agents.Dock = DockStyle.Fill;
            _agents.DropDownStyle = ComboBoxStyle.DropDownList;
            _agents.FlatStyle = FlatStyle.Flat;
            _agents.DisplayMember = "Name";
            _agents.BackColor = _palette.SurfaceRaised;
            _agents.ForeColor = _palette.Text;
            _agents.Font = new Font(Font.FontFamily, 9f, FontStyle.Regular);
            _agents.DrawMode = DrawMode.OwnerDrawFixed;
            _agents.ItemHeight = 24;
            _agents.IntegralHeight = false;
            _agents.DropDownHeight = 240;
            _agents.DrawItem += DrawAgentItem;
            _agents.SelectedIndexChanged += (sender, args) => _sessionId = null;
            var picker = new Panel { Dock = DockStyle.Fill, BackColor = _palette.SurfaceRaised, Margin = new Padding(0) };
            _agents.Dock = DockStyle.Fill;
            var arrow = new RibbonDropArrow(_palette, _agents);
            picker.Controls.Add(_agents);
            picker.Controls.Add(arrow);
            arrow.BringToFront();
            agentRow.Controls.Add(picker, 0, 1);
            _manage.Dock = DockStyle.Fill;
            _manage.Margin = new Padding(6, 0, 0, 0);
            _manage.Click += async (sender, args) => await ManageAgentsAsync();
            agentRow.Controls.Add(_manage, 1, 1);

            layout.Controls.Add(brandRow, 0, 0);
            layout.Controls.Add(agentRow, 0, 1);
            surface.Controls.Add(layout);
            return surface;
        }

        private Control BuildTranscript()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10), Padding = new Padding(12) };
            _transcript.Dock = DockStyle.Fill;
            _transcript.ReadOnly = true;
            _transcript.BorderStyle = BorderStyle.None;
            _transcript.BackColor = _palette.Surface;
            _transcript.ForeColor = _palette.Text;
            _transcriptRegularFont = new Font(Font.FontFamily, 9.25f, FontStyle.Regular);
            _transcriptBoldFont = new Font(Font.FontFamily, 9.25f, FontStyle.Bold);
            _transcript.Font = _transcriptRegularFont;
            _transcript.DetectUrls = true;
            _transcript.HideSelection = false;
            surface.Controls.Add(_transcript);
            return surface;
        }

        private Control BuildComposer()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 8), Padding = new Padding(12, 9, 12, 8) };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = _palette.Surface };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 18));
            layout.Controls.Add(LabelFor("MESSAGE", 7.5f, FontStyle.Bold, _palette.MutedText), 0, 0);
            _prompt.Dock = DockStyle.Fill;
            _prompt.Multiline = true;
            _prompt.BorderStyle = BorderStyle.None;
            _prompt.ScrollBars = ScrollBars.Vertical;
            _prompt.BackColor = _palette.Surface;
            _prompt.ForeColor = _palette.Text;
            _prompt.Font = new Font(Font.FontFamily, 9.5f, FontStyle.Regular);
            _prompt.KeyDown += PromptOnKeyDown;
            layout.Controls.Add(_prompt, 0, 1);
            var hint = LabelFor("Ctrl + Enter to send", 7.5f, FontStyle.Regular, _palette.MutedText);
            hint.TextAlign = ContentAlignment.BottomRight;
            layout.Controls.Add(hint, 0, 2);
            surface.Controls.Add(layout);
            return surface;
        }

        private Control BuildFooter()
        {
            var footer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, BackColor = _palette.Background, Margin = new Padding(0) };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 18));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 82));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 86));
            _statusDot.Anchor = AnchorStyles.Left;
            _status.Margin = new Padding(0);
            _status.Dock = DockStyle.Fill;
            _status.TextAlign = ContentAlignment.MiddleLeft;
            _status.AutoEllipsis = true;
            _status.ForeColor = _palette.MutedText;
            _status.Font = new Font(Font.FontFamily, 8.25f, FontStyle.Regular);
            _cancel.Dock = DockStyle.Fill;
            _cancel.Margin = new Padding(2, 4, 4, 4);
            _cancel.Enabled = false;
            _cancel.Click += async (sender, args) => await CancelAsync();
            _send.Dock = DockStyle.Fill;
            _send.Margin = new Padding(4, 4, 0, 4);
            _send.Click += async (sender, args) => await SendAsync();
            footer.Controls.Add(_statusDot, 0, 0);
            footer.Controls.Add(_status, 1, 0);
            footer.Controls.Add(_cancel, 2, 0);
            footer.Controls.Add(_send, 3, 0);
            return footer;
        }

        private async Task InitializeAsync()
        {
            SetStatus("Connecting to Ribbon Broker…", _palette.MutedText);
            try
            {
                await _runtime.StartAsync();
                var agentCount = await ReloadAgentsAsync();
                SetStatus(agentCount == 0 ? "Install an ACP agent to begin" : "Ready", agentCount == 0 ? _palette.MutedText : _palette.Success);
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private async Task<int> ReloadAgentsAsync()
        {
            string selectedId = null;
            RibbonUiThread.Run(this, () => selectedId = (_agents.SelectedItem as AgentSummary)?.Id);
            var agents = await _runtime.GetInstalledAgentsAsync();
            var orderedAgents = agents.OrderBy(item => item.Name).ToList();
            RibbonUiThread.Run(this, () =>
            {
                _agents.Items.Clear();
                foreach (var agent in orderedAgents) _agents.Items.Add(agent);
                if (_agents.Items.Count > 0)
                {
                    _agents.SelectedItem = _agents.Items.Cast<AgentSummary>().FirstOrDefault(item => item.Id == selectedId) ?? _agents.Items[0];
                }
            });
            return orderedAgents.Count;
        }

        private async Task ManageAgentsAsync()
        {
            using (var dialog = new AgentManagerDialog(_runtime, _palette)) dialog.ShowDialog(this);
            var agentCount = await ReloadAgentsAsync();
            SetStatus(agentCount == 0 ? "Install an ACP agent to begin" : "Ready", agentCount == 0 ? _palette.MutedText : _palette.Success);
        }

        private async Task SendAsync()
        {
            var text = _prompt.Text.Trim();
            var agent = _agents.SelectedItem as AgentSummary;
            if (text.Length == 0) return;
            if (agent == null)
            {
                MessageBox.Show(this, "Install and select an ACP agent first.", "Ribbon", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SetBusy(true);
            AppendTurn(text, agent.Name ?? "Agent");
            _prompt.Clear();
            try
            {
                if (string.IsNullOrWhiteSpace(_sessionId))
                {
                    var session = await _runtime.StartSessionAsync(agent.Id);
                    if (string.IsNullOrWhiteSpace(session.SessionId) && session.AuthenticationMethods != null && session.AuthenticationMethods.Count > 0)
                    {
                        var method = session.AuthenticationMethods[0];
                        SetStatus("Authenticating with " + method.Name + "…", _palette.Accent);
                        await _runtime.AuthenticateAsync(agent.Id, method.Id);
                        session = await _runtime.StartSessionAsync(agent.Id);
                    }
                    _sessionId = session.SessionId;
                    if (string.IsNullOrWhiteSpace(_sessionId)) throw new InvalidOperationException("The ACP agent did not create a session.");
                }
                await _runtime.PromptAsync(_sessionId, text);
                SetStatus("Ready", _palette.Success);
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private async Task CancelAsync()
        {
            if (string.IsNullOrWhiteSpace(_sessionId)) return;
            try
            {
                await _runtime.CancelAsync(_sessionId);
                SetStatus("Cancelling turn…", _palette.Danger);
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private void RuntimeOnSessionUpdate(object sender, SessionUpdateMessage update)
        {
            RibbonUiThread.Post(this, () => HandleSessionUpdate(update));
        }

        private void HandleSessionUpdate(SessionUpdateMessage update)
        {
            if (update.UpdateKind == "agent_message_chunk" && !string.IsNullOrEmpty(update.Text))
            {
                AppendTranscript(update.Text, _palette.Text, FontStyle.Regular);
            }
            else if (update.UpdateKind == "agent_thought_chunk" && !string.IsNullOrEmpty(update.Text))
            {
                SetStatus("Thinking…", _palette.Accent);
            }
            else if (update.UpdateKind == "tool_call" || update.UpdateKind == "tool_call_update")
            {
                SetStatus(string.IsNullOrWhiteSpace(update.ToolName) ? "Using Office tools…" : update.ToolName + " · " + update.Status, _palette.Accent);
            }
            else if (update.UpdateKind == "turn_complete")
            {
                AppendTranscript(Environment.NewLine + Environment.NewLine, _palette.Text, FontStyle.Regular);
            }
        }

        private void PromptOnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && e.Control)
            {
                e.SuppressKeyPress = true;
                _ = SendAsync();
            }
        }

        private void DrawAgentItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            var background = selected ? RibbonDrawing.Blend(_palette.Accent, _palette.SurfaceRaised, 0.70f) : _palette.SurfaceRaised;
            using (var brush = new SolidBrush(background)) e.Graphics.FillRectangle(brush, e.Bounds);
            var agent = _agents.Items[e.Index] as AgentSummary;
            TextRenderer.DrawText(e.Graphics, agent?.Name ?? _agents.Items[e.Index].ToString(), _agents.Font,
                Rectangle.Inflate(e.Bounds, -8, 0), _palette.Text,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
            if ((e.State & DrawItemState.Focus) == DrawItemState.Focus) e.DrawFocusRectangle();
        }

        private void SetBusy(bool busy)
        {
            RibbonUiThread.Run(this, () =>
            {
                _send.Enabled = !busy;
                _cancel.Enabled = busy;
                _agents.Enabled = !busy;
                _manage.Enabled = !busy;
                if (busy) SetStatusCore("Agent is working…", _palette.Accent);
            });
        }

        private void ShowWelcomeMessage()
        {
            _transcript.Clear();
            AppendTranscript("Ready when you are.\n", _palette.Text, FontStyle.Bold);
            AppendTranscript("Ask your agent to inspect, explain, or update the open " + (_runtime.Registration.HostKind ?? "Office") + " document.", _palette.MutedText, FontStyle.Regular);
            _hasConversation = false;
        }

        private void AppendTurn(string text, string agentName)
        {
            if (!_hasConversation)
            {
                _transcript.Clear();
                _hasConversation = true;
            }
            else
            {
                AppendTranscript(Environment.NewLine + Environment.NewLine, _palette.Text, FontStyle.Regular);
            }
            AppendTranscript("YOU\n", _palette.MutedText, FontStyle.Bold);
            AppendTranscript(text + Environment.NewLine + Environment.NewLine, _palette.Text, FontStyle.Regular);
            AppendTranscript(agentName.ToUpperInvariant() + "\n", _palette.Accent, FontStyle.Bold);
        }

        private void AppendTranscript(string text, Color color, FontStyle style)
        {
            _transcript.SelectionStart = _transcript.TextLength;
            _transcript.SelectionLength = 0;
            _transcript.SelectionColor = color;
            _transcript.SelectionFont = style == FontStyle.Bold ? _transcriptBoldFont : _transcriptRegularFont;
            _transcript.AppendText(text);
            _transcript.SelectionStart = _transcript.TextLength;
            _transcript.ScrollToCaret();
        }

        private void SetStatus(string text, Color dotColor)
        {
            RibbonUiThread.Run(this, () => SetStatusCore(text, dotColor));
        }

        private void SetStatusCore(string text, Color dotColor)
        {
            _status.Text = text;
            _statusDot.DotColor = dotColor;
            _statusDot.Invalidate();
        }

        private void ShowError(Exception exception)
        {
            var message = exception.GetBaseException().Message;
            RibbonUiThread.Run(this, () =>
            {
                SetStatusCore(message, _palette.Danger);
                MessageBox.Show(this, message, "Ribbon", MessageBoxButtons.OK, MessageBoxIcon.Error);
            });
        }

        private Label LabelFor(string text, float size, FontStyle style, Color color)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = color,
                BackColor = _palette.Surface,
                Font = new Font(Font.FontFamily, size, style),
                Margin = new Padding(0)
            };
        }
    }
}
