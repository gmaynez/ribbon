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
        private readonly ComboBox _agents = new ComboBox();
        private readonly Button _manage = new Button();
        private readonly RichTextBox _transcript = new RichTextBox();
        private readonly TextBox _prompt = new TextBox();
        private readonly Button _send = new Button();
        private readonly Button _cancel = new Button();
        private readonly Label _status = new Label();
        private string _sessionId;
        private bool _loaded;

        public RibbonSidebarControl(VstoHostRuntime runtime)
        {
            _runtime = runtime ?? throw new ArgumentNullException(nameof(runtime));
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            BuildLayout();
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
            }
            base.Dispose(disposing);
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(10) };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 88));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var header = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2 };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _agents.Dock = DockStyle.Fill;
            _agents.DropDownStyle = ComboBoxStyle.DropDownList;
            _agents.DisplayMember = "Name";
            _agents.SelectedIndexChanged += (sender, args) => _sessionId = null;
            _manage.Text = "Agents…";
            _manage.AutoSize = true;
            _manage.Click += async (sender, args) => await ManageAgentsAsync();
            header.Controls.Add(_agents, 0, 0);
            header.Controls.Add(_manage, 1, 0);

            _transcript.Dock = DockStyle.Fill;
            _transcript.ReadOnly = true;
            _transcript.BorderStyle = BorderStyle.FixedSingle;
            _transcript.BackColor = Color.White;
            _prompt.Dock = DockStyle.Fill;
            _prompt.Multiline = true;
            _prompt.ScrollBars = ScrollBars.Vertical;
            _prompt.KeyDown += PromptOnKeyDown;

            var footer = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3 };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _status.Text = "Connecting…";
            _status.AutoSize = true;
            _status.Padding = new Padding(0, 6, 0, 0);
            _send.Text = "Send";
            _send.AutoSize = true;
            _send.Click += async (sender, args) => await SendAsync();
            _cancel.Text = "Cancel";
            _cancel.AutoSize = true;
            _cancel.Enabled = false;
            _cancel.Click += async (sender, args) => await CancelAsync();
            footer.Controls.Add(_status, 0, 0);
            footer.Controls.Add(_cancel, 1, 0);
            footer.Controls.Add(_send, 2, 0);

            root.Controls.Add(header, 0, 0);
            root.Controls.Add(_transcript, 0, 1);
            root.Controls.Add(_prompt, 0, 2);
            root.Controls.Add(footer, 0, 3);
            Controls.Add(root);
        }

        private async Task InitializeAsync()
        {
            try
            {
                await _runtime.StartAsync();
                await ReloadAgentsAsync();
                _status.Text = _agents.Items.Count == 0 ? "Install an ACP agent to begin." : "Ready";
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private async Task ReloadAgentsAsync()
        {
            var selectedId = (_agents.SelectedItem as AgentSummary)?.Id;
            var agents = await _runtime.GetInstalledAgentsAsync();
            _agents.Items.Clear();
            foreach (var agent in agents.OrderBy(item => item.Name)) _agents.Items.Add(agent);
            if (_agents.Items.Count > 0)
            {
                _agents.SelectedItem = _agents.Items.Cast<AgentSummary>().FirstOrDefault(item => item.Id == selectedId) ?? _agents.Items[0];
            }
        }

        private async Task ManageAgentsAsync()
        {
            using (var dialog = new AgentManagerDialog(_runtime))
            {
                dialog.ShowDialog(this);
            }
            await ReloadAgentsAsync();
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
            Append("You", text);
            _prompt.Clear();
            try
            {
                if (string.IsNullOrWhiteSpace(_sessionId))
                {
                    var session = await _runtime.StartSessionAsync(agent.Id);
                    if (string.IsNullOrWhiteSpace(session.SessionId) && session.AuthenticationMethods != null && session.AuthenticationMethods.Count > 0)
                    {
                        var method = session.AuthenticationMethods[0];
                        _status.Text = "Authenticating with " + method.Name + "…";
                        await _runtime.AuthenticateAsync(agent.Id, method.Id);
                        session = await _runtime.StartSessionAsync(agent.Id);
                    }
                    _sessionId = session.SessionId;
                    if (string.IsNullOrWhiteSpace(_sessionId)) throw new InvalidOperationException("The ACP agent did not create a session.");
                }
                await _runtime.PromptAsync(_sessionId, text);
                _status.Text = "Ready";
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
                _status.Text = "Cancelling…";
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private void RuntimeOnSessionUpdate(object sender, SessionUpdateMessage update)
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action<object, SessionUpdateMessage>(RuntimeOnSessionUpdate), sender, update);
                return;
            }
            if (update.UpdateKind == "agent_message_chunk" && !string.IsNullOrEmpty(update.Text))
            {
                _transcript.AppendText(update.Text);
                _transcript.SelectionStart = _transcript.TextLength;
                _transcript.ScrollToCaret();
            }
            else if (update.UpdateKind == "tool_call" || update.UpdateKind == "tool_call_update")
            {
                _status.Text = string.IsNullOrWhiteSpace(update.ToolName) ? "Using Office tools…" : update.ToolName + " — " + update.Status;
            }
            else if (update.UpdateKind == "turn_complete")
            {
                _transcript.AppendText(Environment.NewLine + Environment.NewLine);
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

        private void SetBusy(bool busy)
        {
            _send.Enabled = !busy;
            _cancel.Enabled = busy;
            _agents.Enabled = !busy;
            _manage.Enabled = !busy;
            if (busy) _status.Text = "Agent is working…";
        }

        private void Append(string speaker, string text)
        {
            if (_transcript.TextLength > 0) _transcript.AppendText(Environment.NewLine + Environment.NewLine);
            _transcript.AppendText(speaker + ": " + text + Environment.NewLine + Environment.NewLine + "Agent: ");
        }

        private void ShowError(Exception exception)
        {
            var message = exception.GetBaseException().Message;
            _status.Text = message;
            MessageBox.Show(this, message, "Ribbon", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
