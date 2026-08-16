using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    internal sealed class SidebarSession
    {
        private readonly VstoHostRuntime _runtime;
        private Task _startTask;
        private string _startAgentId;

        public SidebarSession(VstoHostRuntime runtime)
        {
            _runtime = runtime ?? throw new ArgumentNullException(nameof(runtime));
            ConfigOptions = new List<SessionConfigOption>();
        }

        public string SessionId { get; private set; }
        public string AgentId { get; private set; }
        public string WorkingDirectory { get; private set; }
        public bool SupportsLoad { get; private set; }
        public bool SupportsResume { get; private set; }
        public bool SupportsList { get; private set; }
        public IList<SessionConfigOption> ConfigOptions { get; private set; }
        public bool HasSession => !string.IsNullOrWhiteSpace(SessionId);

        public bool Matches(string agentId)
        {
            return HasSession && string.Equals(AgentId, agentId, StringComparison.OrdinalIgnoreCase);
        }

        public void Reset()
        {
            SessionId = null;
            AgentId = null;
            WorkingDirectory = null;
            SupportsLoad = false;
            SupportsResume = false;
            SupportsList = false;
            ConfigOptions = new List<SessionConfigOption>();
            _runtime.ClearActiveSession();
        }

        public void DropIdentity()
        {
            SessionId = null;
            AgentId = null;
        }

        public void ReplaceConfigOptions(IList<SessionConfigOption> configOptions)
        {
            ConfigOptions = configOptions ?? new List<SessionConfigOption>();
        }

        public async Task CloseAsync()
        {
            var sessionId = SessionId;
            Reset();
            if (!string.IsNullOrWhiteSpace(sessionId)) await _runtime.CloseSessionAsync(sessionId);
        }

        public async Task<bool> EnsureAsync(AgentSummary agent, Action<string> status, Func<bool> accept)
        {
            if (agent == null) return false;
            if (Matches(agent.Id)) return false;
            if (_startTask != null && string.Equals(_startAgentId, agent.Id, StringComparison.OrdinalIgnoreCase))
            {
                await _startTask;
                return Matches(agent.Id);
            }

            var startTask = StartCoreAsync(agent, status, accept);
            _startAgentId = agent.Id;
            _startTask = startTask;
            try
            {
                await startTask;
                return Matches(agent.Id);
            }
            finally
            {
                if (ReferenceEquals(_startTask, startTask))
                {
                    _startTask = null;
                    _startAgentId = null;
                }
            }
        }

        public async Task<SessionResumeResponse> ResumeAsync(AgentSummary agent, ConversationRecord record, Action<string> status)
        {
            if (agent == null) throw new ArgumentNullException(nameof(agent));
            if (record == null) throw new ArgumentNullException(nameof(record));

            status?.Invoke("Restoring " + agent.Name + " context…");
            var resume = await _runtime.ResumeSessionAsync(agent.Id, record.AcpSessionId, record.AcpWorkingDirectory);
            if (!resume.Resumed && HasAuthentication(resume.AuthenticationMethods))
            {
                var method = resume.AuthenticationMethods[0];
                status?.Invoke("Authenticating with " + method.Name + "…");
                await _runtime.AuthenticateAsync(agent.Id, method.Id);
                resume = await _runtime.ResumeSessionAsync(agent.Id, record.AcpSessionId, record.AcpWorkingDirectory);
            }

            if (resume.Resumed) BindResume(agent.Id, resume);
            return resume;
        }

        public Task CancelAsync()
        {
            return HasSession ? _runtime.CancelAsync(SessionId) : Task.CompletedTask;
        }

        public async Task<IList<SessionConfigOption>> SetModelAsync(string configId, string value)
        {
            var response = await _runtime.SetSessionConfigOptionAsync(SessionId, configId, value);
            ReplaceConfigOptions(response.ConfigOptions);
            return ConfigOptions;
        }

        public static SessionConfigOption FindModelOption(IList<SessionConfigOption> configOptions)
        {
            var options = configOptions ?? new List<SessionConfigOption>();
            return options.FirstOrDefault(option =>
                    string.Equals(option.Type, "select", StringComparison.Ordinal)
                    && string.Equals(option.Category, "model", StringComparison.Ordinal))
                ?? options.FirstOrDefault(option =>
                    string.Equals(option.Type, "select", StringComparison.Ordinal)
                    && string.Equals(option.Id, "model", StringComparison.OrdinalIgnoreCase));
        }

        private async Task StartCoreAsync(AgentSummary agent, Action<string> status, Func<bool> accept)
        {
            status?.Invoke("Starting " + agent.Name + "…");
            var session = await _runtime.StartSessionAsync(agent.Id);
            if (string.IsNullOrWhiteSpace(session.SessionId) && HasAuthentication(session.AuthenticationMethods))
            {
                var method = session.AuthenticationMethods[0];
                status?.Invoke("Authenticating with " + method.Name + "…");
                await _runtime.AuthenticateAsync(agent.Id, method.Id);
                session = await _runtime.StartSessionAsync(agent.Id);
            }

            if (string.IsNullOrWhiteSpace(session.SessionId))
            {
                throw new InvalidOperationException("The ACP agent did not create a session.");
            }

            if (accept != null && accept()) BindStart(agent.Id, session);
        }

        private void BindStart(string agentId, SessionStartResponse session)
        {
            SessionId = session.SessionId;
            AgentId = agentId;
            WorkingDirectory = session.WorkingDirectory;
            SupportsLoad = session.SupportsLoad;
            SupportsResume = session.SupportsResume;
            SupportsList = session.SupportsList;
            ReplaceConfigOptions(session.ConfigOptions);
        }

        private void BindResume(string agentId, SessionResumeResponse resume)
        {
            SessionId = resume.SessionId;
            AgentId = agentId;
            WorkingDirectory = resume.WorkingDirectory;
            SupportsLoad = resume.SupportsLoad;
            SupportsResume = resume.SupportsResume;
            SupportsList = resume.SupportsList;
            ReplaceConfigOptions(resume.ConfigOptions);
        }

        private static bool HasAuthentication(IList<AgentAuthenticationMethod> methods)
        {
            return methods != null && methods.Count > 0;
        }
    }
}
