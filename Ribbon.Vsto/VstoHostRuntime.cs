using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    public sealed class VstoHostRuntime : IDisposable
    {
        private readonly IOfficeHost _host;
        private readonly BrokerClient _client;
        private readonly CancellationTokenSource _lifetime = new CancellationTokenSource();
        private Task _startTask;

        public VstoHostRuntime(IOfficeHost host, SynchronizationContext ui)
        {
            _host = host ?? throw new ArgumentNullException(nameof(host));
            _client = new BrokerClient(host, ui ?? throw new ArgumentNullException(nameof(ui)));
            _client.SessionUpdate += (sender, message) => SessionUpdate?.Invoke(this, message);
        }

        public HostRegistration Registration => _host.Registration;
        public event EventHandler<SessionUpdateMessage> SessionUpdate;

        public Task StartAsync()
        {
            return _startTask ?? (_startTask = _client.ConnectAsync(_lifetime.Token));
        }

        public async Task<IList<AgentSummary>> GetInstalledAgentsAsync()
        {
            await StartAsync().ConfigureAwait(false);
            return await _client.RequestAsync<List<AgentSummary>>(RibbonProtocol.ListInstalledAgents, new { }, _lifetime.Token).ConfigureAwait(false);
        }

        public async Task<IList<AgentSummary>> GetRegistryAgentsAsync()
        {
            await StartAsync().ConfigureAwait(false);
            return await _client.RequestAsync<List<AgentSummary>>(RibbonProtocol.ListRegistryAgents, new { }, _lifetime.Token).ConfigureAwait(false);
        }

        public async Task InstallAgentAsync(string agentId)
        {
            await StartAsync().ConfigureAwait(false);
            await _client.RequestAsync(RibbonProtocol.InstallAgent, new AgentIdRequest { AgentId = agentId }, _lifetime.Token).ConfigureAwait(false);
        }

        public async Task UninstallAgentAsync(string agentId)
        {
            await StartAsync().ConfigureAwait(false);
            await _client.RequestAsync(RibbonProtocol.UninstallAgent, new AgentIdRequest { AgentId = agentId }, _lifetime.Token).ConfigureAwait(false);
        }

        public async Task<SessionStartResponse> StartSessionAsync(string agentId)
        {
            await StartAsync().ConfigureAwait(false);
            var documentPath = Registration.DocumentPath;
            var workingDirectory = !string.IsNullOrWhiteSpace(documentPath) ? Path.GetDirectoryName(documentPath) : null;
            return await _client.RequestAsync<SessionStartResponse>(RibbonProtocol.StartSession, new SessionStartRequest
            {
                AgentId = agentId,
                HostId = Registration.HostId,
                WorkingDirectory = workingDirectory
            }, _lifetime.Token).ConfigureAwait(false);
        }

        public async Task AuthenticateAsync(string agentId, string methodId)
        {
            await StartAsync().ConfigureAwait(false);
            await _client.RequestAsync(RibbonProtocol.AuthenticateAgent, new AgentAuthenticationRequest
            {
                AgentId = agentId,
                MethodId = methodId
            }, _lifetime.Token).ConfigureAwait(false);
        }

        public async Task PromptAsync(string sessionId, string text)
        {
            await _client.RequestAsync(RibbonProtocol.PromptSession, new SessionPromptRequest
            {
                SessionId = sessionId,
                Text = text
            }, _lifetime.Token).ConfigureAwait(false);
        }

        public Task CancelAsync(string sessionId)
        {
            return _client.RequestAsync(RibbonProtocol.CancelSession, new SessionCancelRequest { SessionId = sessionId }, _lifetime.Token);
        }

        public void Dispose()
        {
            _lifetime.Cancel();
            _client.Dispose();
            _lifetime.Dispose();
        }
    }
}
