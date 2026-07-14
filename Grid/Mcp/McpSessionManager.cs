using System;

namespace Grid.Mcp
{
    internal sealed class McpSessionManager
    {
        private readonly object _sync = new object();
        private GridMcpSession _activeSession;
        private bool _initializationInProgress;

        public bool TryBeginInitialization()
        {
            lock (_sync)
            {
                if (_initializationInProgress || _activeSession != null)
                {
                    return false;
                }

                _initializationInProgress = true;
                return true;
            }
        }

        public void CompleteInitialization(GridMcpSession session)
        {
            lock (_sync)
            {
                _activeSession = session;
                _initializationInProgress = false;
            }
        }

        public void AbortInitialization()
        {
            lock (_sync)
            {
                _initializationInProgress = false;
            }
        }

        public GridMcpSession GetSession(string sessionId)
        {
            lock (_sync)
            {
                if (_activeSession == null)
                {
                    return null;
                }

                return string.Equals(_activeSession.SessionId, sessionId, StringComparison.Ordinal)
                    ? _activeSession
                    : null;
            }
        }

        public GridMcpSession RemoveSession(string sessionId)
        {
            GridMcpSession session;

            lock (_sync)
            {
                if (_activeSession == null || (sessionId != null && !string.Equals(_activeSession.SessionId, sessionId, StringComparison.Ordinal)))
                {
                    return null;
                }

                session = _activeSession;
                _activeSession = null;
                _initializationInProgress = false;
                return session;
            }
        }

        public GridMcpSession Reset()
        {
            return RemoveSession(null);
        }
    }
}
