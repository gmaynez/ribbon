using System;
using System.Collections.Generic;
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
        private readonly RibbonComboBox _models;
        private readonly RibbonComboBox _checkpoints;
        private readonly RibbonButton _manage;
        private readonly RibbonButton _restoreCheckpoint;
        private readonly RichTextBox _transcript;
        private readonly TextBox _prompt;
        private readonly RibbonButton _send;
        private readonly RibbonButton _cancel;
        private readonly Label _status;
        private readonly RibbonStatusDot _statusDot;
        private readonly Label _promptPlaceholder;
        private readonly ToolTip _toolTip;
        private readonly string _hostKind;
        private readonly string _productName;
        private Font _transcriptRegularFont;
        private Font _transcriptBoldFont;
        private Font _transcriptItalicFont;
        private readonly Dictionary<string, string> _toolStatuses = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly List<DocumentCheckpoint> _checkpointItems = new List<DocumentCheckpoint>();
        private string _sessionId;
        private string _sessionAgentId;
        private string _sessionStartAgentId;
        private string _modelConfigId;
        private Task _sessionStartTask;
        private IList<SessionConfigOption> _sessionConfigOptions = new List<SessionConfigOption>();
        private bool _loaded;
        private bool _hasConversation;
        private bool _suppressAgentSelection;
        private bool _suppressModelSelection;
        private bool _modelAvailable;
        private bool _busy;
        private bool _activityVisible;
        private bool _thoughtVisible;
        private bool _responseVisible;
        private string _planSignature;

        public RibbonSidebarControl(VstoHostRuntime runtime)
        {
            _runtime = runtime ?? throw new ArgumentNullException(nameof(runtime));
            _palette = RibbonPalette.Detect();
            _hostKind = string.IsNullOrWhiteSpace(_runtime.Registration.HostKind) ? "Office" : _runtime.Registration.HostKind;
            _productName = RibbonProductIdentity.GetProductName(_hostKind);
            _agents = new RibbonComboBox(_palette);
            _models = new RibbonComboBox(_palette);
            _checkpoints = new RibbonComboBox(_palette);
            _manage = new RibbonButton(_palette, RibbonButtonKind.Secondary) { Text = "Agents", Glyph = RibbonGlyph.Agents, Width = 88 };
            _restoreCheckpoint = new RibbonButton(_palette, RibbonButtonKind.Secondary) { Text = "Restore", Width = 78 };
            _transcript = new RichTextBox();
            _prompt = new TextBox();
            _send = new RibbonButton(_palette, RibbonButtonKind.Primary) { Text = "Send", Glyph = RibbonGlyph.Send, Width = 80 };
            _cancel = new RibbonButton(_palette, RibbonButtonKind.Ghost) { Text = "Stop", Glyph = RibbonGlyph.Stop, Width = 76 };
            _status = new Label();
            _statusDot = new RibbonStatusDot { DotColor = _palette.MutedText };
            _promptPlaceholder = new Label();
            _toolTip = new ToolTip { InitialDelay = 450, ReshowDelay = 100, AutoPopDelay = 8000 };

            Dock = DockStyle.Fill;
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = _palette.Background;
            ForeColor = _palette.Text;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
            MinimumSize = new Size(300, 360);
            BuildLayout();
            ShowModelPlaceholderCore("Select an agent");
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
                _transcriptItalicFont?.Dispose();
                _toolTip.Dispose();
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
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 178));
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
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = _palette.Surface };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var brandRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, BackColor = _palette.Surface };
            brandRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40));
            brandRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            var mark = new RibbonBrandMark(
                _palette,
                RibbonProductIdentity.GetMark(_hostKind),
                RibbonProductIdentity.GetBrandColor(_hostKind)) { Margin = new Padding(0, 1, 8, 0) };
            var titles = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = _palette.Surface };
            titles.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            titles.RowStyles.Add(new RowStyle(SizeType.Absolute, 18));
            var title = LabelFor("Ribbon " + _productName, 10.5f, FontStyle.Bold, _palette.Text);
            var subtitle = LabelFor("Agent workspace for " + _hostKind, 8f, FontStyle.Regular, _palette.MutedText);
            titles.Controls.Add(title, 0, 0);
            titles.Controls.Add(subtitle, 0, 1);
            brandRow.Controls.Add(mark, 0, 0);
            brandRow.Controls.Add(titles, 1, 0);

            var agentRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, BackColor = _palette.Surface };
            agentRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            agentRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 94));
            agentRow.RowStyles.Add(new RowStyle(SizeType.Absolute, 18));
            agentRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var agentLabel = LabelFor("AGENT", 7.5f, FontStyle.Bold, _palette.MutedText);
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
            _agents.SelectedIndexChanged += async (sender, args) => await AgentSelectionChangedAsync();
            _agents.AccessibleName = "Active ACP agent";
            _toolTip.SetToolTip(_agents, "Choose the ACP agent for this conversation.");
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
            _manage.AccessibleName = "Manage ACP agents";
            _toolTip.SetToolTip(_manage, "Browse, install, update, or remove ACP agents.");
            agentRow.Controls.Add(_manage, 1, 1);

            var modelRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = _palette.Surface, Margin = new Padding(0, 4, 0, 0) };
            modelRow.RowStyles.Add(new RowStyle(SizeType.Absolute, 18));
            modelRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            modelRow.Controls.Add(LabelFor("MODEL", 7.5f, FontStyle.Bold, _palette.MutedText), 0, 0);
            _models.Dock = DockStyle.Fill;
            _models.DropDownStyle = ComboBoxStyle.DropDownList;
            _models.FlatStyle = FlatStyle.Flat;
            _models.DisplayMember = "Name";
            _models.BackColor = _palette.SurfaceRaised;
            _models.ForeColor = _palette.Text;
            _models.Font = new Font(Font.FontFamily, 9f, FontStyle.Regular);
            _models.DrawMode = DrawMode.OwnerDrawFixed;
            _models.ItemHeight = 24;
            _models.IntegralHeight = false;
            _models.DropDownHeight = 280;
            _models.DrawItem += DrawModelItem;
            _models.SelectedIndexChanged += async (sender, args) => await ModelSelectionChangedAsync();
            _models.AccessibleName = "Active model";
            _toolTip.SetToolTip(_models, "Choose a model exposed by the active ACP agent.");
            var modelPicker = new Panel { Dock = DockStyle.Fill, BackColor = _palette.SurfaceRaised, Margin = new Padding(0) };
            var modelArrow = new RibbonDropArrow(_palette, _models);
            modelPicker.Controls.Add(_models);
            modelPicker.Controls.Add(modelArrow);
            modelArrow.BringToFront();
            modelRow.Controls.Add(modelPicker, 0, 1);

            layout.Controls.Add(brandRow, 0, 0);
            layout.Controls.Add(agentRow, 0, 1);
            layout.Controls.Add(modelRow, 0, 2);
            surface.Controls.Add(layout);
            return surface;
        }

        private Control BuildTranscript()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10), Padding = new Padding(12, 9, 12, 12) };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = _palette.Surface };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.Controls.Add(LabelFor("CONVERSATION", 7.5f, FontStyle.Bold, _palette.MutedText), 0, 0);
            _transcript.Dock = DockStyle.Fill;
            _transcript.ReadOnly = true;
            _transcript.BorderStyle = BorderStyle.None;
            _transcript.BackColor = _palette.Surface;
            _transcript.ForeColor = _palette.Text;
            _transcriptRegularFont = new Font(Font.FontFamily, 9.5f, FontStyle.Regular);
            _transcriptBoldFont = new Font(Font.FontFamily, 9.25f, FontStyle.Bold);
            _transcriptItalicFont = new Font(Font.FontFamily, 9.25f, FontStyle.Italic);
            _transcript.Font = _transcriptRegularFont;
            _transcript.DetectUrls = true;
            _transcript.HideSelection = false;
            _transcript.AccessibleName = "Conversation transcript";
            layout.Controls.Add(_transcript, 0, 1);
            layout.Controls.Add(BuildCheckpointBar(), 0, 2);
            surface.Controls.Add(layout);
            return surface;
        }

        private Control BuildCheckpointBar()
        {
            var bar = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = _palette.Surface,
                Margin = new Padding(0, 5, 0, 0)
            };
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 78));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 84));
            bar.Controls.Add(LabelFor("CHECKPOINT", 7.5f, FontStyle.Bold, _palette.MutedText), 0, 0);

            _checkpoints.Dock = DockStyle.Fill;
            _checkpoints.DropDownStyle = ComboBoxStyle.DropDownList;
            _checkpoints.DisplayMember = "DisplayName";
            _checkpoints.Enabled = false;
            _checkpoints.AccessibleName = "Document checkpoint";
            _toolTip.SetToolTip(_checkpoints, "Choose a document state captured before an agent turn.");
            bar.Controls.Add(_checkpoints, 1, 0);

            _restoreCheckpoint.Dock = DockStyle.Fill;
            _restoreCheckpoint.Margin = new Padding(6, 0, 0, 0);
            _restoreCheckpoint.Enabled = false;
            _restoreCheckpoint.Click += async (sender, args) => await RestoreSelectedCheckpointAsync();
            _restoreCheckpoint.AccessibleName = "Restore selected document checkpoint";
            _toolTip.SetToolTip(_restoreCheckpoint, "Restore the open document to the selected checkpoint and start a fresh agent session.");
            bar.Controls.Add(_restoreCheckpoint, 2, 0);
            return bar;
        }

        private Control BuildComposer()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 8), Padding = new Padding(12, 9, 12, 8) };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = _palette.Surface };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 18));
            layout.Controls.Add(LabelFor("PROMPT", 7.5f, FontStyle.Bold, _palette.MutedText), 0, 0);
            var promptHost = new Panel { Dock = DockStyle.Fill, BackColor = _palette.Surface, Margin = new Padding(0) };
            _prompt.Dock = DockStyle.Fill;
            _prompt.Multiline = true;
            _prompt.BorderStyle = BorderStyle.None;
            _prompt.ScrollBars = ScrollBars.None;
            _prompt.WordWrap = true;
            _prompt.AcceptsReturn = true;
            _prompt.BackColor = _palette.Surface;
            _prompt.ForeColor = _palette.Text;
            _prompt.Font = new Font(Font.FontFamily, 9.5f, FontStyle.Regular);
            _prompt.KeyDown += PromptOnKeyDown;
            _prompt.TextChanged += (sender, args) => UpdatePromptPlaceholder();
            _prompt.Enter += (sender, args) => UpdatePromptPlaceholder();
            _prompt.Leave += (sender, args) => UpdatePromptPlaceholder();
            _prompt.AccessibleName = "Prompt for the active agent";
            _promptPlaceholder.AutoSize = true;
            _promptPlaceholder.Text = "Describe what you want to do in " + _hostKind + "…";
            _promptPlaceholder.ForeColor = _palette.MutedText;
            _promptPlaceholder.BackColor = _palette.Surface;
            _promptPlaceholder.Font = new Font(Font.FontFamily, 9.5f, FontStyle.Regular);
            _promptPlaceholder.Location = new Point(0, 2);
            _promptPlaceholder.Cursor = Cursors.IBeam;
            _promptPlaceholder.Click += (sender, args) => _prompt.Focus();
            promptHost.Controls.Add(_prompt);
            promptHost.Controls.Add(_promptPlaceholder);
            _promptPlaceholder.BringToFront();
            layout.Controls.Add(promptHost, 0, 1);
            var hint = LabelFor("Ctrl + Enter to send  ·  Enter for a new line", 7.5f, FontStyle.Regular, _palette.MutedText);
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
            _cancel.AccessibleName = "Stop the active agent turn";
            _send.Dock = DockStyle.Fill;
            _send.Margin = new Padding(4, 4, 0, 4);
            _send.Click += async (sender, args) => await SendAsync();
            _send.AccessibleName = "Send prompt";
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
                if (agentCount > 0) await EnsureSelectedSessionAsync();
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
                _suppressAgentSelection = true;
                try
                {
                    _agents.Items.Clear();
                    foreach (var agent in orderedAgents) _agents.Items.Add(agent);
                    if (_agents.Items.Count > 0)
                    {
                        _agents.SelectedItem = _agents.Items.Cast<AgentSummary>().FirstOrDefault(item => item.Id == selectedId) ?? _agents.Items[0];
                    }
                    else
                    {
                        _sessionId = null;
                        _sessionAgentId = null;
                        ShowModelPlaceholder("Select an agent");
                    }
                }
                finally
                {
                    _suppressAgentSelection = false;
                }
            });
            return orderedAgents.Count;
        }

        private async Task ManageAgentsAsync()
        {
            using (var dialog = new AgentManagerDialog(_runtime, _palette)) dialog.ShowDialog(this);
            var agentCount = await ReloadAgentsAsync();
            if (agentCount > 0) await EnsureSelectedSessionAsync();
            SetStatus(agentCount == 0 ? "Install an ACP agent to begin" : "Ready", agentCount == 0 ? _palette.MutedText : _palette.Success);
        }

        private async Task AgentSelectionChangedAsync()
        {
            if (_suppressAgentSelection) return;
            var previousSessionId = _sessionId;
            _sessionId = null;
            _sessionAgentId = null;
            _sessionConfigOptions = new List<SessionConfigOption>();
            _runtime.ClearActiveSession();
            ShowModelPlaceholder(GetSelectedAgent() == null ? "Select an agent" : "Loading models…");
            try
            {
                if (!string.IsNullOrWhiteSpace(previousSessionId)) await _runtime.CloseSessionAsync(previousSessionId);
                await EnsureSelectedSessionAsync();
                if (GetSelectedAgent() != null) SetStatus("Ready", _palette.Success);
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private Task EnsureSelectedSessionAsync()
        {
            var agent = GetSelectedAgent();
            return agent == null ? Task.CompletedTask : EnsureSessionAsync(agent);
        }

        private async Task EnsureSessionAsync(AgentSummary agent)
        {
            if (!string.IsNullOrWhiteSpace(_sessionId) && string.Equals(_sessionAgentId, agent.Id, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
            if (_sessionStartTask != null && string.Equals(_sessionStartAgentId, agent.Id, StringComparison.OrdinalIgnoreCase))
            {
                await _sessionStartTask;
                return;
            }

            var startTask = StartSessionCoreAsync(agent);
            _sessionStartAgentId = agent.Id;
            _sessionStartTask = startTask;
            try
            {
                await startTask;
            }
            finally
            {
                if (ReferenceEquals(_sessionStartTask, startTask))
                {
                    _sessionStartTask = null;
                    _sessionStartAgentId = null;
                }
            }
        }

        private async Task StartSessionCoreAsync(AgentSummary agent)
        {
            SetStatus("Starting " + agent.Name + "…", _palette.Accent);
            var session = await _runtime.StartSessionAsync(agent.Id);
            if (string.IsNullOrWhiteSpace(session.SessionId) && session.AuthenticationMethods != null && session.AuthenticationMethods.Count > 0)
            {
                var method = session.AuthenticationMethods[0];
                SetStatus("Authenticating with " + method.Name + "…", _palette.Accent);
                await _runtime.AuthenticateAsync(agent.Id, method.Id);
                session = await _runtime.StartSessionAsync(agent.Id);
            }
            if (string.IsNullOrWhiteSpace(session.SessionId))
            {
                throw new InvalidOperationException("The ACP agent did not create a session.");
            }

            var selectedAgent = GetSelectedAgent();
            if (selectedAgent != null && string.Equals(selectedAgent.Id, agent.Id, StringComparison.OrdinalIgnoreCase))
            {
                _sessionId = session.SessionId;
                _sessionAgentId = agent.Id;
                ApplyConfigOptions(session.ConfigOptions);
            }
        }

        private async Task ModelSelectionChangedAsync()
        {
            if (_suppressModelSelection || !_modelAvailable || string.IsNullOrWhiteSpace(_sessionId) || string.IsNullOrWhiteSpace(_modelConfigId))
            {
                return;
            }
            SessionConfigOptionValue selected = null;
            RibbonUiThread.Run(this, () => selected = _models.SelectedItem as SessionConfigOptionValue);
            if (selected == null || string.IsNullOrWhiteSpace(selected.Value)) return;

            SetModelEnabled(false);
            try
            {
                SetStatus("Switching to " + selected.Name + "…", _palette.Accent);
                var response = await _runtime.SetSessionConfigOptionAsync(_sessionId, _modelConfigId, selected.Value);
                ApplyConfigOptions(response.ConfigOptions);
                SetStatus(selected.Name, _palette.Success);
            }
            catch (Exception exception)
            {
                ApplyConfigOptions(_sessionConfigOptions);
                ShowError(exception);
            }
            finally
            {
                SetModelEnabled(!_busy && _modelAvailable);
            }
        }

        private void ApplyConfigOptions(IList<SessionConfigOption> configOptions)
        {
            RibbonUiThread.Run(this, () => ApplyConfigOptionsCore(configOptions));
        }

        private void ApplyConfigOptionsCore(IList<SessionConfigOption> configOptions)
        {
            var options = configOptions ?? new List<SessionConfigOption>();
            _sessionConfigOptions = options.ToList();
            var model = options.FirstOrDefault(option =>
                    string.Equals(option.Type, "select", StringComparison.Ordinal)
                    && string.Equals(option.Category, "model", StringComparison.Ordinal))
                ?? options.FirstOrDefault(option =>
                    string.Equals(option.Type, "select", StringComparison.Ordinal)
                    && string.Equals(option.Id, "model", StringComparison.OrdinalIgnoreCase));
            if (model == null || model.Options == null || model.Options.Count == 0)
            {
                ShowModelPlaceholder("Agent default");
                return;
            }

            _suppressModelSelection = true;
            try
            {
                _models.Items.Clear();
                foreach (var value in model.Options) _models.Items.Add(value);
                var current = _models.Items.Cast<SessionConfigOptionValue>()
                    .FirstOrDefault(value => string.Equals(value.Value, model.CurrentValue, StringComparison.Ordinal));
                if (current == null && !string.IsNullOrWhiteSpace(model.CurrentValue))
                {
                    current = new SessionConfigOptionValue { Value = model.CurrentValue, Name = model.CurrentValue };
                    _models.Items.Insert(0, current);
                }
                _models.SelectedItem = current ?? _models.Items[0];
                _modelConfigId = model.Id;
                _modelAvailable = true;
                _models.Enabled = !_busy;
            }
            finally
            {
                _suppressModelSelection = false;
            }
        }

        private void ShowModelPlaceholder(string text)
        {
            RibbonUiThread.Run(this, () => ShowModelPlaceholderCore(text));
        }

        private void ShowModelPlaceholderCore(string text)
        {
            _suppressModelSelection = true;
            try
            {
                _modelConfigId = null;
                _modelAvailable = false;
                _models.Items.Clear();
                _models.Items.Add(new SessionConfigOptionValue { Value = string.Empty, Name = text });
                _models.SelectedIndex = 0;
                _models.Enabled = false;
            }
            finally
            {
                _suppressModelSelection = false;
            }
        }

        private async Task SendAsync()
        {
            var text = _prompt.Text.Trim();
            var agent = GetSelectedAgent();
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
                await EnsureSessionAsync(agent);
                if (string.IsNullOrWhiteSpace(_sessionId)) throw new InvalidOperationException("The ACP agent did not create a session.");
                SetStatus("Creating document checkpoint…", _palette.Accent);
                var checkpoint = await _runtime.CreateCheckpointAsync(CheckpointLabel(text));
                AddCheckpoint(checkpoint);
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

        private async Task RestoreSelectedCheckpointAsync()
        {
            DocumentCheckpoint selected = null;
            RibbonUiThread.Run(this, () => selected = _checkpoints.SelectedItem as DocumentCheckpoint);
            if (selected == null || _busy) return;

            var answer = MessageBox.Show(
                this,
                "Restore the open " + RibbonProductIdentity.GetDocumentNoun(_hostKind) + " to:\r\n\r\n" + selected.DisplayName
                    + "\r\n\r\nRibbon will first save the current state as another checkpoint. The agent session will restart so it does not rely on stale document context.",
                "Restore Ribbon checkpoint",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.OK) return;

            SetBusy(true);
            try
            {
                SetStatus("Saving current state…", _palette.Accent);
                var safety = await _runtime.CreateCheckpointAsync("Before checkpoint restore");
                SetStatus("Restoring checkpoint…", _palette.Accent);
                await _runtime.RestoreCheckpointAsync(selected);
                AddCheckpoint(safety);

                var previousSessionId = _sessionId;
                _sessionId = null;
                _sessionAgentId = null;
                _sessionConfigOptions = new List<SessionConfigOption>();
                _runtime.ClearActiveSession();
                if (!string.IsNullOrWhiteSpace(previousSessionId)) await _runtime.CloseSessionAsync(previousSessionId);
                ShowModelPlaceholder("Restarting agent session…");
                AppendTranscript(Environment.NewLine + Environment.NewLine + "Checkpoint restored\n", _palette.Success, FontStyle.Bold);
                AppendTranscript("The document is back at " + selected.DisplayName + ". A fresh agent session will be used for the next turn.\n", _palette.MutedText, FontStyle.Regular);
                await EnsureSelectedSessionAsync();
                SetStatus("Checkpoint restored", _palette.Success);
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

        private void AddCheckpoint(DocumentCheckpoint checkpoint)
        {
            if (checkpoint == null) return;
            RibbonUiThread.Run(this, () =>
            {
                _checkpointItems.Insert(0, checkpoint);
                _checkpoints.Items.Insert(0, checkpoint);
                _checkpoints.SelectedIndex = 0;
                while (_checkpointItems.Count > 12)
                {
                    var expired = _checkpointItems[_checkpointItems.Count - 1];
                    _checkpointItems.RemoveAt(_checkpointItems.Count - 1);
                    _checkpoints.Items.Remove(expired);
                    DocumentCheckpointStorage.Delete(expired);
                }
                _checkpoints.Enabled = !_busy && _checkpointItems.Count > 0;
                _restoreCheckpoint.Enabled = !_busy && _checkpointItems.Count > 0;
            });
        }

        private static string CheckpointLabel(string prompt)
        {
            var singleLine = string.Join(" ", (prompt ?? string.Empty)
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                .Trim();
            if (singleLine.Length > 42) singleLine = singleLine.Substring(0, 39) + "…";
            return string.IsNullOrWhiteSpace(singleLine) ? "Before agent turn" : "Before: " + singleLine;
        }

        private void RuntimeOnSessionUpdate(object sender, SessionUpdateMessage update)
        {
            RibbonUiThread.Post(this, () => HandleSessionUpdate(update));
        }

        private void HandleSessionUpdate(SessionUpdateMessage update)
        {
            if (update.UpdateKind == "agent_message_chunk" && !string.IsNullOrEmpty(update.Text))
            {
                if (_activityVisible && !_responseVisible)
                {
                    AppendTranscript(Environment.NewLine + "Response\n", _palette.Accent, FontStyle.Bold);
                    _responseVisible = true;
                }
                AppendTranscript(update.Text, _palette.Text, FontStyle.Regular);
            }
            else if (update.UpdateKind == "agent_thought_chunk" && !string.IsNullOrEmpty(update.Text))
            {
                if (!_thoughtVisible)
                {
                    AppendTranscript("Thinking\n", _palette.MutedText, FontStyle.Bold);
                    _thoughtVisible = true;
                    _activityVisible = true;
                }
                AppendTranscript(update.Text, _palette.MutedText, FontStyle.Italic);
                SetStatus("Thinking…", _palette.Accent);
            }
            else if (update.UpdateKind == "tool_call" || update.UpdateKind == "tool_call_update")
            {
                AppendToolActivity(update);
                var status = FriendlyToolStatus(update.Status);
                SetStatus(string.IsNullOrWhiteSpace(update.ToolName)
                    ? "Using Office tools…"
                    : update.ToolName + (string.IsNullOrWhiteSpace(status) ? string.Empty : " · " + status),
                    string.Equals(update.Status, "failed", StringComparison.OrdinalIgnoreCase) ? _palette.Danger : _palette.Accent);
            }
            else if (update.UpdateKind == "plan")
            {
                AppendPlan(update.PlanEntries);
            }
            else if (update.UpdateKind == "config_option_update")
            {
                ApplyConfigOptions(update.ConfigOptions);
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

        private void DrawModelItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            var background = selected ? RibbonDrawing.Blend(_palette.Accent, _palette.SurfaceRaised, 0.70f) : _palette.SurfaceRaised;
            using (var brush = new SolidBrush(background)) e.Graphics.FillRectangle(brush, e.Bounds);
            var model = _models.Items[e.Index] as SessionConfigOptionValue;
            TextRenderer.DrawText(e.Graphics, model?.Name ?? _models.Items[e.Index].ToString(), _models.Font,
                Rectangle.Inflate(e.Bounds, -8, 0), _palette.Text,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
            if ((e.State & DrawItemState.Focus) == DrawItemState.Focus) e.DrawFocusRectangle();
        }

        private void SetBusy(bool busy)
        {
            RibbonUiThread.Run(this, () =>
            {
                _busy = busy;
                _send.Enabled = !busy;
                _cancel.Enabled = busy;
                _agents.Enabled = !busy;
                _models.Enabled = !busy && _modelAvailable;
                _manage.Enabled = !busy;
                _checkpoints.Enabled = !busy && _checkpointItems.Count > 0;
                _restoreCheckpoint.Enabled = !busy && _checkpointItems.Count > 0;
                if (busy) SetStatusCore("Agent is working…", _palette.Accent);
            });
        }

        private AgentSummary GetSelectedAgent()
        {
            AgentSummary selected = null;
            RibbonUiThread.Run(this, () => selected = _agents.SelectedItem as AgentSummary);
            return selected;
        }

        private void SetModelEnabled(bool enabled)
        {
            RibbonUiThread.Run(this, () => _models.Enabled = enabled);
        }

        private void ShowWelcomeMessage()
        {
            _transcript.Clear();
            AppendTranscript("Ready to work in " + _hostKind + ".\n", _palette.Text, FontStyle.Bold);
            AppendTranscript(
                "Ask an agent to inspect, explain, or update the open " + RibbonProductIdentity.GetDocumentNoun(_hostKind) + ".\n\n",
                _palette.MutedText,
                FontStyle.Regular);
            AppendTranscript(RibbonProductIdentity.GetExamplePrompt(_hostKind), _palette.Accent, FontStyle.Regular);
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
            AppendTranscript("You\n", _palette.MutedText, FontStyle.Bold);
            AppendTranscript(text + Environment.NewLine + Environment.NewLine, _palette.Text, FontStyle.Regular);
            AppendTranscript(agentName + "\n", _palette.Accent, FontStyle.Bold);
            _toolStatuses.Clear();
            _activityVisible = false;
            _thoughtVisible = false;
            _responseVisible = false;
            _planSignature = null;
        }

        private void AppendPlan(IList<SessionPlanEntry> entries)
        {
            if (entries == null || entries.Count == 0) return;
            var signature = string.Join("|", entries.Select(entry =>
                (entry.Content ?? string.Empty) + ":" + (entry.Status ?? string.Empty)));
            if (string.Equals(signature, _planSignature, StringComparison.Ordinal)) return;
            _planSignature = signature;
            AppendTranscript((_activityVisible ? Environment.NewLine : string.Empty) + "Plan\n", _palette.MutedText, FontStyle.Bold);
            foreach (var entry in entries)
            {
                var marker = string.Equals(entry.Status, "completed", StringComparison.OrdinalIgnoreCase) ? "✓"
                    : string.Equals(entry.Status, "in_progress", StringComparison.OrdinalIgnoreCase) ? "→"
                    : "•";
                var color = marker == "✓" ? _palette.Success : marker == "→" ? _palette.Accent : _palette.MutedText;
                AppendTranscript(marker + " " + (entry.Content ?? string.Empty) + "\n", color, FontStyle.Regular);
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
                AppendTranscript("Activity\n", _palette.MutedText, FontStyle.Bold);
            }

            if (first)
            {
                AppendTranscript("• " + (update.ToolName ?? "Office action") + "\n", _palette.Text, FontStyle.Regular);
                _activityVisible = true;
            }

            if (!string.Equals(previousStatus, status, StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals(status, "completed", StringComparison.OrdinalIgnoreCase))
                {
                    AppendTranscript("  ✓ Completed\n", _palette.Success, FontStyle.Regular);
                }
                else if (string.Equals(status, "failed", StringComparison.OrdinalIgnoreCase))
                {
                    AppendTranscript("  × Failed\n", _palette.Danger, FontStyle.Regular);
                }
                else if (!first && string.Equals(status, "in_progress", StringComparison.OrdinalIgnoreCase))
                {
                    AppendTranscript("  → In progress\n", _palette.Accent, FontStyle.Regular);
                }
                _toolStatuses[key] = status;
            }

            if (!string.IsNullOrWhiteSpace(update.Text))
            {
                AppendTranscript("  " + update.Text.Trim() + "\n", _palette.MutedText, FontStyle.Regular);
            }
        }

        private static string FriendlyToolStatus(string status)
        {
            if (string.IsNullOrWhiteSpace(status)) return string.Empty;
            return status.Replace('_', ' ');
        }

        private void UpdatePromptPlaceholder()
        {
            RibbonUiThread.Run(this, () =>
                _promptPlaceholder.Visible = _prompt.TextLength == 0 && !_prompt.Focused);
        }

        private void AppendTranscript(string text, Color color, FontStyle style)
        {
            _transcript.SelectionStart = _transcript.TextLength;
            _transcript.SelectionLength = 0;
            _transcript.SelectionColor = color;
            _transcript.SelectionFont = style == FontStyle.Bold
                ? _transcriptBoldFont
                : style == FontStyle.Italic ? _transcriptItalicFont : _transcriptRegularFont;
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
