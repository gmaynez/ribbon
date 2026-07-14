using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    internal sealed class AgentManagerDialog : Form
    {
        private readonly VstoHostRuntime _runtime;
        private readonly ListView _list = new ListView();
        private readonly Button _toggle = new Button();
        private readonly Button _refresh = new Button();
        private readonly Label _status = new Label();

        public AgentManagerDialog(VstoHostRuntime runtime)
        {
            _runtime = runtime;
            Text = "ACP Agent Registry";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(680, 420);
            Size = new Size(760, 500);
            BuildLayout();
        }

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
            await ReloadAsync();
        }

        private void BuildLayout()
        {
            _list.Dock = DockStyle.Fill;
            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.HideSelection = false;
            _list.Columns.Add("Agent", 190);
            _list.Columns.Add("Version", 90);
            _list.Columns.Add("Distribution", 90);
            _list.Columns.Add("Status", 100);
            _list.Columns.Add("Description", 260);
            _list.SelectedIndexChanged += (sender, args) => UpdateToggle();

            var footer = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(8) };
            var close = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.OK };
            _toggle.Text = "Install";
            _toggle.AutoSize = true;
            _toggle.Enabled = false;
            _toggle.Click += async (sender, args) => await ToggleAsync();
            _refresh.Text = "Refresh";
            _refresh.AutoSize = true;
            _refresh.Click += async (sender, args) => await ReloadAsync();
            _status.AutoSize = true;
            _status.Padding = new Padding(0, 6, 12, 0);
            footer.Controls.Add(close);
            footer.Controls.Add(_toggle);
            footer.Controls.Add(_refresh);
            footer.Controls.Add(_status);
            Controls.Add(_list);
            Controls.Add(footer);
            AcceptButton = close;
        }

        private async Task ReloadAsync()
        {
            SetBusy(true, "Loading registry…");
            try
            {
                var agents = await _runtime.GetRegistryAgentsAsync();
                _list.Items.Clear();
                foreach (var agent in agents)
                {
                    var row = new ListViewItem(agent.Name ?? agent.Id) { Tag = agent };
                    row.SubItems.Add(agent.Version ?? string.Empty);
                    row.SubItems.Add(agent.DistributionType ?? string.Empty);
                    row.SubItems.Add(agent.Installed ? (agent.UpdateAvailable ? "Update available" : "Installed") : "Available");
                    row.SubItems.Add(agent.Description ?? string.Empty);
                    _list.Items.Add(row);
                }
                _status.Text = agents.Count + " compatible agents";
            }
            catch (Exception exception)
            {
                _status.Text = exception.GetBaseException().Message;
                MessageBox.Show(this, _status.Text, "ACP Agent Registry", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false, _status.Text);
                UpdateToggle();
            }
        }

        private async Task ToggleAsync()
        {
            if (_list.SelectedItems.Count == 0) return;
            var agent = _list.SelectedItems[0].Tag as AgentSummary;
            if (agent == null) return;
            SetBusy(true, agent.Installed ? "Uninstalling…" : "Installing…");
            try
            {
                if (agent.Installed) await _runtime.UninstallAgentAsync(agent.Id);
                else await _runtime.InstallAgentAsync(agent.Id);
                await ReloadAsync();
            }
            catch (Exception exception)
            {
                _status.Text = exception.GetBaseException().Message;
                MessageBox.Show(this, _status.Text, "ACP Agent Registry", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetBusy(false, _status.Text);
            }
        }

        private void UpdateToggle()
        {
            var agent = _list.SelectedItems.Count == 0 ? null : _list.SelectedItems[0].Tag as AgentSummary;
            _toggle.Enabled = agent != null;
            _toggle.Text = agent != null && agent.Installed ? "Uninstall" : "Install";
        }

        private void SetBusy(bool busy, string status)
        {
            _list.Enabled = !busy;
            _toggle.Enabled = !busy && _list.SelectedItems.Count > 0;
            _refresh.Enabled = !busy;
            _status.Text = status;
            UseWaitCursor = busy;
        }
    }
}
