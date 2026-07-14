using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Grid.Chat;
using Grid.Runtime;

namespace Grid.Sidebar
{
    internal sealed class GridSidebarControl : UserControl
    {
        private readonly GridRuntime _runtime;
        private readonly RichTextBox _transcriptBox;
        private readonly TextBox _inputBox;
        private readonly TextBox _baseUrlBox;
        private readonly TextBox _modelBox;
        private readonly TextBox _apiKeyBox;
        private readonly Label _statusLabel;
        private readonly Button _sendButton;
        private readonly Button _saveButton;
        private readonly Button _clearButton;

        public GridSidebarControl(GridRuntime runtime)
        {
            TableLayoutPanel root;
            TableLayoutPanel settingsLayout;
            TableLayoutPanel inputLayout;

            _runtime = runtime ?? throw new ArgumentNullException(nameof(runtime));
            Dock = DockStyle.Fill;
            Padding = new Padding(8);

            root = new TableLayoutPanel();
            root.ColumnCount = 1;
            root.RowCount = 3;
            root.Dock = DockStyle.Fill;
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            settingsLayout = new TableLayoutPanel();
            settingsLayout.ColumnCount = 2;
            settingsLayout.RowCount = 5;
            settingsLayout.Dock = DockStyle.Top;
            settingsLayout.AutoSize = true;
            settingsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90F));
            settingsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            _baseUrlBox = CreateTextBox();
            _modelBox = CreateTextBox();
            _apiKeyBox = CreateTextBox();
            _apiKeyBox.UseSystemPasswordChar = true;
            _statusLabel = new Label();
            _statusLabel.AutoSize = true;
            _statusLabel.Text = "Ready.";

            _saveButton = new Button();
            _saveButton.Text = "Save Settings";
            _saveButton.AutoSize = true;
            _saveButton.Click += SaveButton_Click;

            settingsLayout.Controls.Add(CreateLabel("Base URL"), 0, 0);
            settingsLayout.Controls.Add(_baseUrlBox, 1, 0);
            settingsLayout.Controls.Add(CreateLabel("Model"), 0, 1);
            settingsLayout.Controls.Add(_modelBox, 1, 1);
            settingsLayout.Controls.Add(CreateLabel("API Key"), 0, 2);
            settingsLayout.Controls.Add(_apiKeyBox, 1, 2);
            settingsLayout.Controls.Add(CreateLabel("MCP"), 0, 3);
            settingsLayout.Controls.Add(CreateLabel(string.IsNullOrWhiteSpace(_runtime.McpServerHost.McpEndpointUrl) ? "Not running" : _runtime.McpServerHost.McpEndpointUrl), 1, 3);
            settingsLayout.Controls.Add(_saveButton, 1, 4);

            _transcriptBox = new RichTextBox();
            _transcriptBox.Dock = DockStyle.Fill;
            _transcriptBox.ReadOnly = true;
            _transcriptBox.BackColor = SystemColors.Window;
            _transcriptBox.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            inputLayout = new TableLayoutPanel();
            inputLayout.ColumnCount = 2;
            inputLayout.RowCount = 3;
            inputLayout.Dock = DockStyle.Bottom;
            inputLayout.AutoSize = true;
            inputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            inputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _inputBox = new TextBox();
            _inputBox.Dock = DockStyle.Fill;
            _inputBox.Multiline = true;
            _inputBox.Height = 84;
            _inputBox.ScrollBars = ScrollBars.Vertical;

            _sendButton = new Button();
            _sendButton.Text = "Send";
            _sendButton.Width = 92;
            _sendButton.Height = 32;
            _sendButton.Click += SendButton_Click;

            _clearButton = new Button();
            _clearButton.Text = "Clear";
            _clearButton.Width = 92;
            _clearButton.Height = 32;
            _clearButton.Click += ClearButton_Click;

            inputLayout.Controls.Add(_inputBox, 0, 0);
            inputLayout.SetRowSpan(_inputBox, 2);
            inputLayout.Controls.Add(_sendButton, 1, 0);
            inputLayout.Controls.Add(_clearButton, 1, 1);
            inputLayout.Controls.Add(_statusLabel, 0, 2);
            inputLayout.SetColumnSpan(_statusLabel, 2);

            root.Controls.Add(settingsLayout, 0, 0);
            root.Controls.Add(_transcriptBox, 0, 1);
            root.Controls.Add(inputLayout, 0, 2);

            Controls.Add(root);
            LoadSettings();
            AppendTranscript("System", "Grid sidebar ready. Configure your provider and start chatting.");
        }

        private void LoadSettings()
        {
            _baseUrlBox.Text = _runtime.Settings.ProviderBaseUrl;
            _modelBox.Text = _runtime.Settings.ProviderModel;
            _apiKeyBox.Text = _runtime.Settings.GetProviderApiKey();
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            SaveSettings();
        }

        private async void SendButton_Click(object sender, EventArgs e)
        {
            string userMessage;
            ConversationTurnResult result;

            userMessage = _inputBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(userMessage))
            {
                return;
            }

            try
            {
                SaveSettings();
                SetBusyState(true, "Working...");
                AppendTranscript("User", userMessage);
                _inputBox.Clear();

                result = await _runtime.ConversationService.SendAsync(userMessage, CancellationToken.None).ConfigureAwait(true);
                if (result.ExecutedTools.Count > 0)
                {
                    AppendTranscript("Tools", string.Join(", ", result.ExecutedTools.Distinct().ToArray()));
                }

                AppendTranscript("Assistant", string.IsNullOrWhiteSpace(result.AssistantMessage) ? "(No text response)" : result.AssistantMessage);
                SetBusyState(false, "Ready.");
            }
            catch (Exception ex)
            {
                AppendTranscript("Error", ex.Message);
                SetBusyState(false, "Request failed.");
            }
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            _runtime.ConversationService.ClearConversation();
            _transcriptBox.Clear();
            AppendTranscript("System", "Conversation cleared.");
        }

        private void SaveSettings()
        {
            _runtime.Settings.ProviderBaseUrl = _baseUrlBox.Text.Trim();
            _runtime.Settings.ProviderModel = _modelBox.Text.Trim();
            _runtime.Settings.SetProviderApiKey(_apiKeyBox.Text);
            _runtime.Settings.SaveNormalized();
            _statusLabel.Text = "Settings saved.";
        }

        private void SetBusyState(bool isBusy, string statusText)
        {
            _sendButton.Enabled = !isBusy;
            _saveButton.Enabled = !isBusy;
            _clearButton.Enabled = !isBusy;
            _inputBox.Enabled = !isBusy;
            _statusLabel.Text = statusText;
        }

        private void AppendTranscript(string speaker, string text)
        {
            if (_transcriptBox.TextLength > 0)
            {
                _transcriptBox.AppendText(Environment.NewLine + Environment.NewLine);
            }

            _transcriptBox.SelectionFont = new Font(_transcriptBox.Font, FontStyle.Bold);
            _transcriptBox.AppendText(speaker + ": ");
            _transcriptBox.SelectionFont = new Font(_transcriptBox.Font, FontStyle.Regular);
            _transcriptBox.AppendText(text ?? string.Empty);
            _transcriptBox.SelectionStart = _transcriptBox.TextLength;
            _transcriptBox.ScrollToCaret();
        }

        private static Label CreateLabel(string text)
        {
            Label label;

            label = new Label();
            label.AutoSize = true;
            label.Margin = new Padding(0, 6, 6, 6);
            label.Text = text;
            return label;
        }

        private static TextBox CreateTextBox()
        {
            TextBox textBox;

            textBox = new TextBox();
            textBox.Dock = DockStyle.Top;
            textBox.Margin = new Padding(0, 2, 0, 6);
            return textBox;
        }
    }
}
