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
        private readonly SidebarSession _session;
        private readonly ConversationWorkspace _workspace;
        private readonly TranscriptView _transcriptView;
        private readonly SidebarCheckpoints _checkpointList;
        private readonly RibbonPalette _palette;
        private readonly RibbonComboBox _agents;
        private readonly RibbonComboBox _models;
        private readonly RibbonComboBox _checkpoints;
        private readonly RibbonButton _manage;
        private readonly RibbonButton _newConversation;
        private readonly RibbonButton _history;
        private readonly RibbonButton _restoreCheckpoint;
        private readonly RibbonButton _approvalModeToggle;
        private readonly RichTextBox _transcript;
        private readonly TextBox _prompt;
        private readonly RibbonButton _send;
        private readonly RibbonButton _cancel;
        private readonly Label _status;
        private readonly RibbonStatusDot _statusDot;
        private readonly Label _promptPlaceholder;
        private readonly ToolTip _toolTip;
        private readonly Timer _historySaveTimer;
        private readonly string _hostKind;
        private readonly string _productName;
        private TableLayoutPanel _rootLayout;
        private TableLayoutPanel _agentRow;
        private TableLayoutPanel _conversationHeader;
        private TableLayoutPanel _checkpointBar;
        private TableLayoutPanel _footer;
        private Label _checkpointLabel;
        private Label _composerHint;
        private string _modelConfigId;
        private bool _loaded;
        private bool _suppressAgentSelection;
        private bool _suppressModelSelection;
        private bool _modelAvailable;
        private bool _busy;

        public RibbonSidebarControl(VstoHostRuntime runtime)
        {
            _runtime = runtime ?? throw new ArgumentNullException(nameof(runtime));
            _session = new SidebarSession(_runtime);
            _workspace = new ConversationWorkspace(_runtime, ScheduleConversationSave);
            _checkpointList = new SidebarCheckpoints();
            _palette = RibbonPalette.Detect();
            _hostKind = string.IsNullOrWhiteSpace(_runtime.Registration.HostKind) ? "Office" : _runtime.Registration.HostKind;
            _productName = RibbonProductIdentity.GetProductName(_hostKind);
            _agents = new RibbonComboBox(_palette);
            _models = new RibbonComboBox(_palette);
            _checkpoints = new RibbonComboBox(_palette);
            _agents.PlaceholderText = "No agents installed";
            _models.PlaceholderText = "Select an agent";
            _checkpoints.PlaceholderText = "No checkpoints yet";
            _manage = new RibbonButton(_palette, RibbonButtonKind.Secondary) { Text = "Agents", Glyph = RibbonGlyph.Agents, Width = 88 };
            _newConversation = new RibbonButton(_palette, RibbonButtonKind.Ghost) { Text = "New", Width = 60, Height = 26, MinimumSize = new Size(56, 26) };
            _history = new RibbonButton(_palette, RibbonButtonKind.Ghost) { Text = "History", Width = 72, Height = 26, MinimumSize = new Size(68, 26) };
            _restoreCheckpoint = new RibbonButton(_palette, RibbonButtonKind.Secondary) { Text = "Restore", Width = 78 };
            _approvalModeToggle = new RibbonButton(_palette, RibbonButtonKind.Ghost) { Text = "Ask", Width = 62, Height = 26, MinimumSize = new Size(54, 26) };
            _transcript = new RichTextBox();
            _transcriptView = new TranscriptView(_transcript, _palette, _hostKind);
            _transcriptView.EntryCaptured += _workspace.Capture;
            _prompt = new TextBox();
            _send = new RibbonButton(_palette, RibbonButtonKind.Primary) { Text = "Send", Glyph = RibbonGlyph.Send, Width = 80 };
            _cancel = new RibbonButton(_palette, RibbonButtonKind.Ghost) { Text = "Stop", Glyph = RibbonGlyph.Stop, Width = 76 };
            _status = new Label();
            _statusDot = new RibbonStatusDot { DotColor = _palette.MutedText };
            _promptPlaceholder = new Label();
            _toolTip = new ToolTip { InitialDelay = 450, ReshowDelay = 100, AutoPopDelay = 8000 };
            _historySaveTimer = new Timer { Interval = 700 };
            _historySaveTimer.Tick += (sender, args) =>
            {
                _historySaveTimer.Stop();
                _workspace.Persist();
            };

            Dock = DockStyle.Fill;
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = _palette.Background;
            ForeColor = _palette.Text;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
            MinimumSize = new Size(300, 360);
            BuildLayout();
            ApplyResponsiveLayout();
            ShowModelPlaceholderCore("Select an agent");
            _transcriptView.ShowWelcome();
            _runtime.SessionUpdate += RuntimeOnSessionUpdate;
            _runtime.ApprovalModeChanged += RuntimeOnApprovalModeChanged;
            _runtime.AutoApproved += RuntimeOnAutoApproved;
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
                _runtime.ApprovalModeChanged -= RuntimeOnApprovalModeChanged;
                _runtime.AutoApproved -= RuntimeOnAutoApproved;
                _historySaveTimer.Stop();
                _workspace.Persist();
                _historySaveTimer.Dispose();
                _transcriptView.Dispose();
                _toolTip.Dispose();
            }
            base.Dispose(disposing);
        }

        private void BuildLayout()
        {
            _rootLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(12),
                BackColor = _palette.Background
            };
            _rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 184));
            _rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 116));
            _rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));

            _rootLayout.Controls.Add(BuildHeader(), 0, 0);
            _rootLayout.Controls.Add(BuildTranscript(), 0, 1);
            _rootLayout.Controls.Add(BuildComposer(), 0, 2);
            _rootLayout.Controls.Add(BuildFooter(), 0, 3);
            Controls.Add(_rootLayout);
        }

        private Control BuildHeader()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10), Padding = new Padding(12, 10, 12, 10) };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = _palette.Surface };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var brandRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, BackColor = _palette.Surface };
            brandRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40));
            brandRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            brandRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var mark = new RibbonBrandMark(
                _palette,
                RibbonProductIdentity.GetMark(_hostKind),
                RibbonProductIdentity.GetBrandColor(_hostKind)) { Margin = new Padding(0, 1, 8, 0) };
            var titles = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, BackColor = _palette.Surface };
            titles.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            titles.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var title = LabelFor("Ribbon " + _productName, 11f, FontStyle.Bold, _palette.Text);
            var subtitle = LabelFor("Agent workspace for " + _hostKind, 8.25f, FontStyle.Regular, _palette.MutedText);
            titles.Controls.Add(title, 0, 0);
            titles.Controls.Add(subtitle, 0, 1);
            brandRow.Controls.Add(mark, 0, 0);
            brandRow.Controls.Add(titles, 1, 0);

            _agentRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, BackColor = _palette.Surface };
            _agentRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _agentRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 94));
            _agentRow.RowStyles.Add(new RowStyle(SizeType.Absolute, 18));
            _agentRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var agentLabel = LabelFor("AGENT", 7.5f, FontStyle.Bold, _palette.MutedText);
            _agentRow.Controls.Add(agentLabel, 0, 0);
            _agentRow.SetColumnSpan(agentLabel, 2);
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
            _agentRow.Controls.Add(picker, 0, 1);
            _manage.Dock = DockStyle.Fill;
            _manage.Margin = new Padding(6, 0, 0, 0);
            _manage.Click += async (sender, args) => await ManageAgentsAsync();
            _manage.AccessibleName = "Manage ACP agents";
            _toolTip.SetToolTip(_manage, "Browse, install, update, or remove ACP agents.");
            _agentRow.Controls.Add(_manage, 1, 1);

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
            layout.Controls.Add(_agentRow, 0, 1);
            layout.Controls.Add(modelRow, 0, 2);
            surface.Controls.Add(layout);
            return surface;
        }

        private Control BuildTranscript()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10), Padding = new Padding(12, 9, 12, 12) };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = _palette.Surface };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.Controls.Add(BuildConversationHeader(), 0, 0);
            _transcript.Dock = DockStyle.Fill;
            _transcript.ReadOnly = true;
            _transcript.BorderStyle = BorderStyle.None;
            _transcript.BackColor = _palette.Surface;
            _transcript.ForeColor = _palette.Text;
            _transcriptView.CreateFonts(Font);
            _transcript.DetectUrls = true;
            _transcript.HideSelection = false;
            _transcript.AccessibleName = "Conversation transcript";
            RibbonNativeTheme.ApplyDarkScrollBars(_transcript, _palette);
            layout.Controls.Add(_transcript, 0, 1);
            layout.Controls.Add(BuildCheckpointBar(), 0, 2);
            surface.Controls.Add(layout);
            return surface;
        }

        private Control BuildConversationHeader()
        {
            _conversationHeader = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, BackColor = _palette.Surface, Margin = new Padding(0) };
            _conversationHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _conversationHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 64));
            _conversationHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 76));
            _conversationHeader.Controls.Add(LabelFor("CONVERSATION", 7.5f, FontStyle.Bold, _palette.MutedText), 0, 0);
            _newConversation.Dock = DockStyle.Fill;
            _newConversation.Margin = new Padding(2, 0, 2, 0);
            _newConversation.Click += async (sender, args) => await NewConversationAsync();
            _newConversation.AccessibleName = "Start a new Ribbon conversation";
            _history.Dock = DockStyle.Fill;
            _history.Margin = new Padding(2, 0, 0, 0);
            _history.Click += async (sender, args) => await ShowHistoryAsync();
            _history.AccessibleName = "Open Ribbon conversation history";
            _toolTip.SetToolTip(_newConversation, "Save this chat and start a fresh agent conversation.");
            _toolTip.SetToolTip(_history, "Browse saved Ribbon conversations.");
            _conversationHeader.Controls.Add(_newConversation, 1, 0);
            _conversationHeader.Controls.Add(_history, 2, 0);
            return _conversationHeader;
        }

        private Control BuildCheckpointBar()
        {
            _checkpointBar = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = _palette.Surface,
                Margin = new Padding(0, 5, 0, 0)
            };
            _checkpointBar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 78));
            _checkpointBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _checkpointBar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 84));
            _checkpointLabel = LabelFor("CHECKPOINT", 7.5f, FontStyle.Bold, _palette.MutedText);
            _checkpointBar.Controls.Add(_checkpointLabel, 0, 0);

            _checkpoints.Dock = DockStyle.Fill;
            _checkpoints.DropDownStyle = ComboBoxStyle.DropDownList;
            _checkpoints.DisplayMember = "DisplayName";
            _checkpoints.Enabled = false;
            _checkpoints.AccessibleName = "Document checkpoint";
            _toolTip.SetToolTip(_checkpoints, "Choose a document state captured before an agent turn.");
            _checkpointBar.Controls.Add(_checkpoints, 1, 0);

            _restoreCheckpoint.Dock = DockStyle.Fill;
            _restoreCheckpoint.Margin = new Padding(6, 0, 0, 0);
            _restoreCheckpoint.Enabled = false;
            _restoreCheckpoint.Click += async (sender, args) => await RestoreSelectedCheckpointAsync();
            _restoreCheckpoint.AccessibleName = "Restore selected document checkpoint";
            _toolTip.SetToolTip(_restoreCheckpoint, "Restore the open document to the selected checkpoint and start a fresh agent session.");
            _checkpointBar.Controls.Add(_restoreCheckpoint, 2, 0);
            return _checkpointBar;
        }

        private Control BuildComposer()
        {
            var surface = new RibbonSurface(_palette) { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 8), Padding = new Padding(12, 9, 12, 8) };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = _palette.Surface };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 18));
            layout.Controls.Add(LabelFor("PROMPT", 7.5f, FontStyle.Bold, _palette.MutedText), 0, 0);
            var promptHost = new RibbonSurface(_palette)
            {
                Dock = DockStyle.Fill,
                BackColor = _palette.SurfaceRaised,
                Margin = new Padding(0, 0, 0, 2),
                Padding = new Padding(10, 6, 10, 6),
                CornerRadius = 7,
                UseRaisedBackground = true
            };
            _prompt.Dock = DockStyle.Fill;
            _prompt.Multiline = true;
            _prompt.BorderStyle = BorderStyle.None;
            _prompt.ScrollBars = ScrollBars.None;
            _prompt.WordWrap = true;
            _prompt.AcceptsReturn = true;
            _prompt.BackColor = _palette.SurfaceRaised;
            _prompt.ForeColor = _palette.Text;
            _prompt.Font = new Font(Font.FontFamily, 9.5f, FontStyle.Regular);
            _prompt.KeyDown += PromptOnKeyDown;
            _prompt.TextChanged += (sender, args) => UpdatePromptPlaceholder();
            _prompt.Enter += (sender, args) => { promptHost.EmphasizeBorder = true; UpdatePromptPlaceholder(); };
            _prompt.Leave += (sender, args) => { promptHost.EmphasizeBorder = false; UpdatePromptPlaceholder(); };
            _prompt.AccessibleName = "Prompt for the active agent";
            _promptPlaceholder.AutoSize = true;
            _promptPlaceholder.Text = "Describe what you want to do in " + _hostKind + "…";
            _promptPlaceholder.ForeColor = _palette.MutedText;
            _promptPlaceholder.BackColor = _palette.SurfaceRaised;
            _promptPlaceholder.Font = new Font(Font.FontFamily, 9.5f, FontStyle.Regular);
            _promptPlaceholder.Location = new Point(10, 8);
            _promptPlaceholder.Cursor = Cursors.IBeam;
            _promptPlaceholder.Click += (sender, args) => _prompt.Focus();
            promptHost.Controls.Add(_prompt);
            promptHost.Controls.Add(_promptPlaceholder);
            _promptPlaceholder.BringToFront();
            layout.Controls.Add(promptHost, 0, 1);
            _composerHint = LabelFor("Ctrl + Enter to send  ·  Enter for a new line", 7.5f, FontStyle.Regular, _palette.MutedText);
            _composerHint.TextAlign = ContentAlignment.BottomRight;
            layout.Controls.Add(_composerHint, 0, 2);
            surface.Controls.Add(layout);
            return surface;
        }

        private Control BuildFooter()
        {
            _footer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, BackColor = _palette.Background, Margin = new Padding(0) };
            _footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 18));
            _footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 62));
            _footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 82));
            _footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 86));
            _statusDot.Anchor = AnchorStyles.Left;
            _status.Margin = new Padding(0);
            _status.Dock = DockStyle.Fill;
            _status.TextAlign = ContentAlignment.MiddleLeft;
            _status.AutoEllipsis = true;
            _status.ForeColor = _palette.MutedText;
            _status.Font = new Font(Font.FontFamily, 8.25f, FontStyle.Regular);
            _approvalModeToggle.Dock = DockStyle.Fill;
            _approvalModeToggle.Margin = new Padding(2, 4, 2, 4);
            _approvalModeToggle.Enabled = false;
            _approvalModeToggle.Click += (sender, args) => CycleApprovalMode();
            _approvalModeToggle.AccessibleName = "Approval mode";
            _toolTip.SetToolTip(_approvalModeToggle, "Ask before each document change. Click to turn on Auto-approve for this agent session.");
            _cancel.Dock = DockStyle.Fill;
            _cancel.Margin = new Padding(2, 4, 4, 4);
            _cancel.Enabled = false;
            _cancel.Click += async (sender, args) => await CancelAsync();
            _cancel.AccessibleName = "Stop the active agent turn";
            _send.Dock = DockStyle.Fill;
            _send.Margin = new Padding(4, 4, 0, 4);
            _send.Click += async (sender, args) => await SendAsync();
            _send.AccessibleName = "Send prompt";
            _footer.Controls.Add(_statusDot, 0, 0);
            _footer.Controls.Add(_status, 1, 0);
            _footer.Controls.Add(_approvalModeToggle, 2, 0);
            _footer.Controls.Add(_cancel, 3, 0);
            _footer.Controls.Add(_send, 4, 0);
            return _footer;
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            ApplyResponsiveLayout();
        }

        private void ApplyResponsiveLayout()
        {
            if (_rootLayout == null) return;
            var compact = ClientSize.Width > 0 && ClientSize.Width < 380;
            _rootLayout.Padding = compact ? new Padding(8) : new Padding(12);
            if (_agentRow != null) _agentRow.ColumnStyles[1].Width = compact ? 84 : 94;
            if (_conversationHeader != null)
            {
                _conversationHeader.ColumnStyles[1].Width = compact ? 54 : 64;
                _conversationHeader.ColumnStyles[2].Width = compact ? 66 : 76;
            }
            if (_checkpointBar != null)
            {
                _checkpointBar.ColumnStyles[0].Width = compact ? 0 : 78;
                _checkpointBar.ColumnStyles[2].Width = compact ? 76 : 84;
            }
            if (_checkpointLabel != null) _checkpointLabel.Visible = !compact;
            if (_footer != null)
            {
                _footer.ColumnStyles[2].Width = compact ? 52 : 62;
            }
            if (_composerHint != null)
            {
                _composerHint.Text = compact ? "Ctrl + Enter to send" : "Ctrl + Enter to send  ·  Enter for a new line";
            }
            if (_footer != null)
            {
                _footer.ColumnStyles[2].Width = compact ? 72 : 82;
                _footer.ColumnStyles[3].Width = compact ? 78 : 86;
            }
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
                        _session.DropIdentity();
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
            _workspace.Persist();
            _workspace.Clear();
            _transcriptView.ShowWelcome();
            ShowModelPlaceholder(GetSelectedAgent() == null ? "Select an agent" : "Loading models…");
            try
            {
                await _session.CloseAsync();
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
            if (await _session.EnsureAsync(agent, message => SetStatus(message, _palette.Accent), () => AgentStillSelected(agent)))
            {
                ApplyBoundSession();
            }
        }

        private void ApplyBoundSession()
        {
            ApplyConfigOptions(_session.ConfigOptions);
            _workspace.ApplySession(_session, GetSelectedModelName());
            RibbonUiThread.Run(this, () => _approvalModeToggle.Enabled = !_busy && _session.HasSession);
        }

        private bool AgentStillSelected(AgentSummary agent)
        {
            var selected = GetSelectedAgent();
            return selected != null && string.Equals(selected.Id, agent.Id, StringComparison.OrdinalIgnoreCase);
        }

        private async Task ModelSelectionChangedAsync()
        {
            if (_suppressModelSelection || !_modelAvailable || !_session.HasSession || string.IsNullOrWhiteSpace(_modelConfigId))
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
                var options = await _session.SetModelAsync(_modelConfigId, selected.Value);
                ApplyConfigOptions(options);
                if (_workspace.HasCurrent)
                {
                    _workspace.ApplyModelName(selected.Name ?? selected.Value);
                }
                SetStatus(selected.Name, _palette.Success);
            }
            catch (Exception exception)
            {
                ApplyConfigOptions(_session.ConfigOptions);
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
            _session.ReplaceConfigOptions(configOptions);
            var model = SidebarSession.FindModelOption(configOptions);
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
            if (_workspace.HistoryReadOnly)
            {
                MessageBox.Show(this, "This conversation belongs to another document and is open read-only. Start a new chat to continue in the current document.",
                    "Ribbon history", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (agent == null)
            {
                MessageBox.Show(this, "Install and select an ACP agent first.", "Ribbon", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (_workspace.HasCurrent && !_workspace.MatchesCurrentDocument())
            {
                var answer = MessageBox.Show(this,
                    "The active " + RibbonProductIdentity.GetDocumentNoun(_hostKind) + " has changed since this conversation began.\r\n\r\nStart a new conversation for the active document?",
                    "Ribbon document changed", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                if (answer != DialogResult.Yes) return;
                await NewConversationAsync();
                agent = GetSelectedAgent();
                if (agent == null) return;
            }

            SetBusy(true);
            _workspace.Ensure(agent, text, _session);
            _transcriptView.BeginTurn(text, agent.Name ?? "Agent");
            _prompt.Clear();
            try
            {
                await EnsureSessionAsync(agent);
                if (!_session.HasSession) throw new InvalidOperationException("The ACP agent did not create a session.");
                _workspace.ApplySession(_session, GetSelectedModelName());
                SetStatus("Creating document checkpoint…", _palette.Accent);
                var checkpoint = await _runtime.CreateCheckpointAsync(SidebarCheckpoints.LabelFor(text));
                AddCheckpoint(checkpoint);
                await _runtime.PromptAsync(_session.SessionId, text);
                _workspace.Persist();
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
            if (!_session.HasSession) return;
            try
            {
                await _session.CancelAsync();
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

                await _session.CloseAsync();
                ShowModelPlaceholder("Restarting agent session…");
                RibbonUiThread.Run(this, () => _transcriptView.AppendCheckpointRestored(selected.DisplayName));
                await EnsureSelectedSessionAsync();
                _workspace.Persist();
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
                var expired = _checkpointList.Add(checkpoint);
                _checkpoints.Items.Insert(0, checkpoint);
                _checkpoints.SelectedIndex = 0;
                foreach (var item in expired) _checkpoints.Items.Remove(item);
                _checkpoints.Enabled = !_busy && _checkpointList.Count > 0;
                _restoreCheckpoint.Enabled = !_busy && _checkpointList.Count > 0;
            });
        }

        private void RuntimeOnSessionUpdate(object sender, SessionUpdateMessage update)
        {
            RibbonUiThread.Post(this, () => HandleSessionUpdate(update));
        }

        private void RuntimeOnApprovalModeChanged(object sender, ApprovalMode mode)
        {
            RibbonUiThread.Post(this, () => ApplyApprovalModeCore(mode));
        }

        private void RuntimeOnAutoApproved(object sender, AutoApprovalRecord record)
        {
            RibbonUiThread.Post(this, () => AppendAutoApproval(record));
        }

        private void ApplyApprovalModeCore(ApprovalMode mode)
        {
            _approvalModeToggle.Text = mode == ApprovalMode.Auto ? "Auto" : "Ask";
            _approvalModeToggle.Kind = mode == ApprovalMode.Auto ? RibbonButtonKind.Danger : RibbonButtonKind.Ghost;
            _toolTip.SetToolTip(_approvalModeToggle, mode == ApprovalMode.Auto
                ? "Auto-approve is on for this session. Click to turn it off and ask before each change."
                : "Ask before each document change. Click to turn on Auto-approve for this agent session.");
        }

        private void AppendAutoApproval(AutoApprovalRecord record)
        {
            var action = record.Category == "acp" ? record.Action : FriendlyApprovalAction(record.Action);
            var label = string.IsNullOrWhiteSpace(action) ? "document change" : action;
            _transcriptView.AppendAutoApproval(label);
        }

        private static string FriendlyApprovalAction(string action)
        {
            if (string.IsNullOrWhiteSpace(action)) return "document change";
            return action
                .Replace("excel_", string.Empty)
                .Replace("word_", string.Empty)
                .Replace("powerpoint_", string.Empty)
                .Replace('_', ' ');
        }

        private void CycleApprovalMode()
        {
            if (_busy) return;
            var next = _runtime.ApprovalMode == ApprovalMode.Auto ? ApprovalMode.Ask : ApprovalMode.Auto;
            if (next == ApprovalMode.Auto)
            {
                var answer = MessageBox.Show(
                    this,
                    "Turn on Auto-approve for this agent session?\r\n\r\n" +
                    "Ribbon will allow every document-changing Office action and agent permission request without asking. " +
                    "Each approved action is logged in this transcript, and you can still restore a checkpoint from the bar above.\r\n\r\n" +
                    "Auto-approve turns off automatically when you start a new conversation, switch agents, or restart " + _hostKind + ".",
                    "Ribbon · Auto-approve",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);
                if (answer != DialogResult.OK) return;
            }
            _runtime.SetApprovalMode(next);
        }

        private void HandleSessionUpdate(SessionUpdateMessage update)
        {
            if (update.UpdateKind == "config_option_update")
            {
                ApplyConfigOptions(update.ConfigOptions);
                return;
            }

            if (update.UpdateKind == "session_info_update")
            {
                _workspace.ApplyTitle(update.Title);
                return;
            }

            if (update.UpdateKind == "turn_complete")
            {
                _transcriptView.EndTurn();
                _workspace.Persist();
                return;
            }

            if (!_transcriptView.TryApply(update)) return;
            if (update.UpdateKind == "agent_thought_chunk")
            {
                SetStatus("Thinking…", _palette.Accent);
            }
            else if (update.UpdateKind == "tool_call" || update.UpdateKind == "tool_call_update")
            {
                SetStatus(TranscriptView.ToolStatusText(update),
                    TranscriptView.IsFailedTool(update) ? _palette.Danger : _palette.Accent);
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
                _send.Enabled = !busy && !_workspace.HistoryReadOnly;
                _cancel.Enabled = busy;
                _prompt.Enabled = !busy && !_workspace.HistoryReadOnly;
                _agents.Enabled = !busy && !_workspace.HistoryReadOnly;
                _models.Enabled = !busy && !_workspace.HistoryReadOnly && _modelAvailable;
                _manage.Enabled = !busy;
                _newConversation.Enabled = !busy;
                _history.Enabled = !busy;
                _checkpoints.Enabled = !busy && _checkpointList.Count > 0;
                _restoreCheckpoint.Enabled = !busy && _checkpointList.Count > 0;
                _approvalModeToggle.Enabled = !busy && _session.HasSession;
                if (busy) SetStatusCore("Agent is working…", _palette.Accent);
            });
        }

        private AgentSummary GetSelectedAgent()
        {
            AgentSummary selected = null;
            RibbonUiThread.Run(this, () => selected = _agents.SelectedItem as AgentSummary);
            return selected;
        }

        private string GetSelectedModelName()
        {
            SessionConfigOptionValue model = null;
            RibbonUiThread.Run(this, () => model = _models.SelectedItem as SessionConfigOptionValue);
            return model != null && !string.IsNullOrWhiteSpace(model.Value) ? model.Name ?? model.Value : null;
        }

        private void SetModelEnabled(bool enabled)
        {
            RibbonUiThread.Run(this, () => _models.Enabled = enabled);
        }

        private async Task NewConversationAsync()
        {
            if (_busy) return;
            SetBusy(true);
            try
            {
                _workspace.Persist();
                _workspace.Clear();
                await _session.CloseAsync();
                _transcriptView.ShowWelcome();
                ShowModelPlaceholder(GetSelectedAgent() == null ? "Select an agent" : "Starting fresh session…");
                await EnsureSelectedSessionAsync();
                SetStatus("New conversation", _palette.Success);
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

        private async Task ShowHistoryAsync()
        {
            if (_busy) return;
            _workspace.Persist();
            ConversationRecord selected;
            using (var dialog = new ConversationHistoryDialog(_runtime.Registration, _palette))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                selected = dialog.SelectedConversation;
            }
            if (selected != null) await OpenConversationAsync(selected);
        }

        private async Task OpenConversationAsync(ConversationRecord record)
        {
            if (record == null || _busy) return;
            if (_workspace.IsOpen(record))
            {
                SetStatus("Conversation is already open", _palette.Success);
                return;
            }

            SetBusy(true);
            try
            {
                _workspace.Persist();
                _workspace.Clear();
                await _session.CloseAsync();

                _transcriptView.Render(record);
                var agent = SelectAgent(record.AgentId);
                if (!_workspace.MatchesDocument(record))
                {
                    OpenReadOnlyHistory(record, "This conversation belongs to " + record.DisplayDocument + ". Open that document to continue with its agent context.");
                    return;
                }

                if (agent == null)
                {
                    OpenReadOnlyHistory(record, "The " + (record.AgentName ?? record.AgentId) + " agent is not installed. The transcript is available read-only.");
                    return;
                }

                if (!record.MayResumeNatively)
                {
                    await OfferFreshContinuationAsync(record, agent, "This agent did not advertise ACP session load or resume support when the conversation was saved.");
                    return;
                }

                var restore = await _workspace.TryRestoreAsync(record, agent, _session, message => SetStatus(message, _palette.Accent));
                if (restore.Kind == ConversationRestoreKind.UnknownSession)
                {
                    await OfferFreshContinuationAsync(record, agent, "The agent no longer lists this saved ACP session.");
                    return;
                }

                if (restore.Kind == ConversationRestoreKind.Resumed)
                {
                    _workspace.Adopt(record);
                    ApplyBoundSession();
                    SetStatus(restore.ResumeKind == "loaded" ? "Conversation loaded" : "Conversation resumed", _palette.Success);
                    return;
                }

                await OfferFreshContinuationAsync(record, agent, restore.Error);
            }
            catch (Exception exception)
            {
                OpenReadOnlyHistory(record, exception.GetBaseException().Message);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private async Task OfferFreshContinuationAsync(ConversationRecord source, AgentSummary agent, string reason)
        {
            var answer = MessageBox.Show(this,
                (string.IsNullOrWhiteSpace(reason) ? "The agent could not restore its previous ACP context." : reason)
                    + "\r\n\r\nContinue in a fresh agent session? The previous transcript will remain visible, but the agent will not remember it unless you summarize the relevant details in your next prompt.",
                "Continue Ribbon conversation", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.Yes)
            {
                OpenReadOnlyHistory(source, "Saved transcript · agent context was not resumed");
                return;
            }

            var continuation = _workspace.CreateFreshContinuation(source, agent);
            _workspace.Adopt(continuation);
            _transcriptView.Render(continuation);
            ShowModelPlaceholder("Starting fresh session…");
            await EnsureSessionAsync(agent);
            _transcriptView.AppendFreshContinuationNotice();
            _workspace.Persist();
            SetStatus("Fresh continuation ready", _palette.Success);
        }

        private void OpenReadOnlyHistory(ConversationRecord record, string status)
        {
            _workspace.OpenReadOnly();
            _transcriptView.Render(record);
            ShowModelPlaceholder("Read-only history");
            SetStatus(status, _palette.MutedText);
        }

        private AgentSummary SelectAgent(string agentId)
        {
            var agent = _agents.Items.Cast<AgentSummary>()
                .FirstOrDefault(item => string.Equals(item.Id, agentId, StringComparison.OrdinalIgnoreCase));
            if (agent == null) return null;
            _suppressAgentSelection = true;
            try { _agents.SelectedItem = agent; }
            finally { _suppressAgentSelection = false; }
            return agent;
        }

        private void ScheduleConversationSave()
        {
            RibbonUiThread.Run(this, () =>
            {
                if (!_workspace.HasCurrent || IsDisposed || Disposing) return;
                _historySaveTimer.Stop();
                _historySaveTimer.Start();
            });
        }

        private void UpdatePromptPlaceholder()
        {
            RibbonUiThread.Run(this, () =>
                _promptPlaceholder.Visible = _prompt.TextLength == 0 && !_prompt.Focused);
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
