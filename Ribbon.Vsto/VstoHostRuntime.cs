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
        private readonly string _hostId;
        private readonly CancellationTokenSource _lifetime = new CancellationTokenSource();
        private readonly CancellationToken _lifetimeToken;
        private Task _startTask;
        private int _disposed;

        public VstoHostRuntime(IOfficeHost host, SynchronizationContext ui)
        {
            _host = host ?? throw new ArgumentNullException(nameof(host));
            _hostId = host.Registration.HostId;
            _client = new BrokerClient(host, ui ?? throw new ArgumentNullException(nameof(ui)));
            _lifetimeToken = _lifetime.Token;
            _client.SessionUpdate += (sender, message) => SessionUpdate?.Invoke(this, message);
            _client.ApprovalModeChanged += (sender, mode) => ApprovalModeChanged?.Invoke(this, mode);
            _client.AutoApproved += (sender, record) => AutoApproved?.Invoke(this, record);
        }

        public HostRegistration Registration => _host.Registration;
        public bool SupportsCheckpoints => _host is ICheckpointHost;
        public event EventHandler<SessionUpdateMessage> SessionUpdate;
        internal event EventHandler<ApprovalMode> ApprovalModeChanged;
        internal event EventHandler<AutoApprovalRecord> AutoApproved;

        internal ApprovalMode ApprovalMode => _client.ApprovalMode;

        internal void SetApprovalMode(ApprovalMode mode) => _client.SetApprovalMode(mode);

        public Task StartAsync()
        {
            if (Volatile.Read(ref _disposed) != 0)
            {
                return Task.FromCanceled(_lifetimeToken.IsCancellationRequested ? _lifetimeToken : new CancellationToken(true));
            }
            return _startTask ?? (_startTask = _client.ConnectAsync(_lifetimeToken));
        }

        public async Task<IList<AgentSummary>> GetInstalledAgentsAsync()
        {
            await StartAsync().ConfigureAwait(false);
            return await _client.RequestAsync<List<AgentSummary>>(RibbonProtocol.ListInstalledAgents, new { }, _lifetimeToken).ConfigureAwait(false);
        }

        public async Task<IList<AgentSummary>> GetRegistryAgentsAsync()
        {
            await StartAsync().ConfigureAwait(false);
            return await _client.RequestAsync<List<AgentSummary>>(RibbonProtocol.ListRegistryAgents, new { }, _lifetimeToken).ConfigureAwait(false);
        }

        public async Task InstallAgentAsync(string agentId)
        {
            await StartAsync().ConfigureAwait(false);
            await _client.RequestAsync(RibbonProtocol.InstallAgent, new AgentIdRequest { AgentId = agentId }, _lifetimeToken).ConfigureAwait(false);
        }

        public async Task UninstallAgentAsync(string agentId)
        {
            await StartAsync().ConfigureAwait(false);
            await _client.RequestAsync(RibbonProtocol.UninstallAgent, new AgentIdRequest { AgentId = agentId }, _lifetimeToken).ConfigureAwait(false);
        }

        public async Task<SessionStartResponse> StartSessionAsync(string agentId)
        {
            await StartAsync().ConfigureAwait(false);
            var documentPath = Registration.DocumentPath;
            var workingDirectory = !string.IsNullOrWhiteSpace(documentPath) ? Path.GetDirectoryName(documentPath) : null;
            var response = await _client.RequestAsync<SessionStartResponse>(RibbonProtocol.StartSession, new SessionStartRequest
            {
                AgentId = agentId,
                HostId = Registration.HostId,
                WorkingDirectory = workingDirectory
            }, _lifetimeToken).ConfigureAwait(false);
            if (!string.IsNullOrWhiteSpace(response.SessionId)) _client.SetActiveSession(response.SessionId);
            return response;
        }

        public async Task<SessionResumeResponse> ResumeSessionAsync(string agentId, string sessionId, string workingDirectory)
        {
            await StartAsync().ConfigureAwait(false);
            var response = await _client.RequestAsync<SessionResumeResponse>(RibbonProtocol.ResumeSession, new SessionResumeRequest
            {
                AgentId = agentId,
                HostId = Registration.HostId,
                SessionId = sessionId,
                WorkingDirectory = workingDirectory
            }, _lifetimeToken).ConfigureAwait(false);
            if (response.Resumed && !string.IsNullOrWhiteSpace(response.SessionId)) _client.SetActiveSession(response.SessionId);
            return response;
        }

        public async Task<AgentSessionListResponse> ListAgentSessionsAsync(string agentId, string workingDirectory)
        {
            await StartAsync().ConfigureAwait(false);
            return await _client.RequestAsync<AgentSessionListResponse>(RibbonProtocol.ListAgentSessions, new AgentSessionListRequest
            {
                AgentId = agentId,
                WorkingDirectory = workingDirectory
            }, _lifetimeToken).ConfigureAwait(false);
        }

        public void ClearActiveSession()
        {
            _client.SetActiveSession(string.Empty);
        }

        public async Task AuthenticateAsync(string agentId, string methodId)
        {
            await StartAsync().ConfigureAwait(false);
            await _client.RequestAsync(RibbonProtocol.AuthenticateAgent, new AgentAuthenticationRequest
            {
                AgentId = agentId,
                MethodId = methodId
            }, _lifetimeToken).ConfigureAwait(false);
        }

        public async Task PromptAsync(string sessionId, string text)
        {
            await _client.RequestAsync(RibbonProtocol.PromptSession, new SessionPromptRequest
            {
                SessionId = sessionId,
                Text = text
            }, _lifetimeToken).ConfigureAwait(false);
        }

        public Task CancelAsync(string sessionId)
        {
            return _client.RequestAsync(RibbonProtocol.CancelSession, new SessionCancelRequest { SessionId = sessionId }, _lifetimeToken);
        }

        public Task CloseSessionAsync(string sessionId)
        {
            if (string.IsNullOrWhiteSpace(sessionId)) return Task.CompletedTask;
            return _client.RequestAsync(RibbonProtocol.CloseSession, new SessionCancelRequest { SessionId = sessionId }, _lifetimeToken);
        }

        public Task<SessionConfigOptionsResponse> SetSessionConfigOptionAsync(string sessionId, string configId, string value)
        {
            return _client.RequestAsync<SessionConfigOptionsResponse>(RibbonProtocol.SetSessionConfigOption, new SessionConfigOptionRequest
            {
                SessionId = sessionId,
                ConfigId = configId,
                Value = value
            }, _lifetimeToken);
        }

        public Task<DocumentCheckpoint> CreateCheckpointAsync(string label)
        {
            var checkpointHost = _host as ICheckpointHost;
            if (checkpointHost == null)
            {
                throw new NotSupportedException("This Office host does not support document checkpoints.");
            }
            return checkpointHost.CreateCheckpointAsync(label, _lifetimeToken);
        }

        public Task RestoreCheckpointAsync(DocumentCheckpoint checkpoint)
        {
            var checkpointHost = _host as ICheckpointHost;
            if (checkpointHost == null)
            {
                throw new NotSupportedException("This Office host does not support document checkpoints.");
            }
            return checkpointHost.RestoreCheckpointAsync(checkpoint, _lifetimeToken);
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
            _lifetime.Cancel();
            _client.Dispose();
            DocumentCheckpointStorage.DeleteHostCheckpoints(_hostId);
            // Async task-pane continuations may still hold the token while Office tears down.
            // Let the process reclaim the source instead of racing those continuations.
        }

    }
}
