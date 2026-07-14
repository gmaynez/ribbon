using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using Grid.Tools;

namespace Grid.Mcp
{
    internal sealed class GridMcpServerHost : IDisposable
    {
        private readonly GridToolCatalog _toolCatalog;
        private readonly LocalRequestValidator _requestValidator;
        private readonly McpSessionManager _sessionManager;

        private HttpListener _listener;
        private CancellationTokenSource _listenerCancellationSource;
        private Task _listenLoopTask;
        private bool _disposed;

        public GridMcpServerHost(GridToolCatalog toolCatalog)
        {
            _toolCatalog = toolCatalog ?? throw new ArgumentNullException(nameof(toolCatalog));
            _requestValidator = new LocalRequestValidator();
            _sessionManager = new McpSessionManager();
        }

        public int Port { get; private set; }

        public bool IsRunning
        {
            get { return _listener != null && _listener.IsListening; }
        }

        public string McpEndpointUrl
        {
            get { return Port > 0 ? string.Format("http://localhost:{0}/mcp", Port) : string.Empty; }
        }

        public Task StartAsync(int port, CancellationToken cancellationToken)
        {
            string prefix;

            if (IsRunning)
            {
                return Task.CompletedTask;
            }

            prefix = string.Format("http://localhost:{0}/", port);
            _listener = new HttpListener();
            _listener.Prefixes.Add(prefix);
            _listenerCancellationSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            _listener.Start();
            Port = port;
            _listenLoopTask = ListenLoopAsync(_listenerCancellationSource.Token);
            return Task.CompletedTask;
        }

        public async Task StopAsync()
        {
            GridMcpSession session;

            session = _sessionManager.Reset();
            if (_listenerCancellationSource != null)
            {
                try
                {
                    _listenerCancellationSource.Cancel();
                }
                catch
                {
                }
            }

            if (_listener != null)
            {
                try
                {
                    _listener.Stop();
                }
                catch
                {
                }
            }

            if (_listenLoopTask != null)
            {
                try
                {
                    await _listenLoopTask.ConfigureAwait(false);
                }
                catch
                {
                }
            }

            if (session != null)
            {
                await session.DisposeAsync().ConfigureAwait(false);
            }

            if (_listener != null)
            {
                try
                {
                    _listener.Close();
                }
                catch
                {
                }
            }

            if (_listenerCancellationSource != null)
            {
                _listenerCancellationSource.Dispose();
            }

            _listener = null;
            _listenerCancellationSource = null;
            _listenLoopTask = null;
            Port = 0;
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            StopAsync().GetAwaiter().GetResult();
        }

        private async Task ListenLoopAsync(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested && _listener != null && _listener.IsListening)
            {
                HttpListenerContext context;

                try
                {
                    context = await _listener.GetContextAsync().ConfigureAwait(false);
                }
                catch (HttpListenerException)
                {
                    break;
                }
                catch (ObjectDisposedException)
                {
                    break;
                }

                _ = Task.Run(delegate { return HandleContextAsync(context, cancellationToken); }, cancellationToken);
            }
        }

        private async Task HandleContextAsync(HttpListenerContext context, CancellationToken cancellationToken)
        {
            HttpListenerRequest request;
            HttpListenerResponse response;

            request = context.Request;
            response = context.Response;

            try
            {
                if (!_requestValidator.Validate(request, response))
                {
                    return;
                }

                if (string.Equals(request.Url.AbsolutePath, "/health", StringComparison.OrdinalIgnoreCase))
                {
                    WriteJsonResponse(response, 200, "{\"status\":\"ready\",\"transport\":\"streamable-http\"}");
                    return;
                }

                if (!string.Equals(request.Url.AbsolutePath, "/mcp", StringComparison.OrdinalIgnoreCase))
                {
                    WriteJsonResponse(response, 404, "{\"error\":\"Not found\"}");
                    return;
                }

                if (string.Equals(request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase))
                {
                    await HandlePostAsync(request, response, cancellationToken).ConfigureAwait(false);
                    return;
                }

                if (string.Equals(request.HttpMethod, "GET", StringComparison.OrdinalIgnoreCase))
                {
                    await HandleGetAsync(request, response, cancellationToken).ConfigureAwait(false);
                    return;
                }

                if (string.Equals(request.HttpMethod, "DELETE", StringComparison.OrdinalIgnoreCase))
                {
                    await HandleDeleteAsync(request, response).ConfigureAwait(false);
                    return;
                }

                WriteJsonResponse(response, 405, "{\"error\":\"Method not allowed\"}");
            }
            catch (JsonException ex)
            {
                WriteJsonResponse(response, 400, JsonSerializer.Serialize(new { error = ex.Message }));
            }
            catch (Exception ex)
            {
                WriteJsonResponse(response, 500, JsonSerializer.Serialize(new { error = ex.Message }));
            }
            finally
            {
                try
                {
                    response.Close();
                }
                catch
                {
                }
            }
        }

        private async Task HandlePostAsync(HttpListenerRequest request, HttpListenerResponse response, CancellationToken cancellationToken)
        {
            JsonRpcMessage message;
            string body;
            string sessionId;
            bool wroteResponse;
            GridMcpSession session;

            using (StreamReader reader = new StreamReader(request.InputStream, request.ContentEncoding))
            {
                body = await reader.ReadToEndAsync().ConfigureAwait(false);
            }

            message = JsonSerializer.Deserialize<JsonRpcMessage>(body, McpJsonUtilities.DefaultOptions);
            if (message == null)
            {
                throw new JsonException("Request body did not contain a valid JSON-RPC message.");
            }

            sessionId = request.Headers["Mcp-Session-Id"];
            if (string.IsNullOrWhiteSpace(sessionId))
            {
                await HandleInitializePostAsync(message, response, cancellationToken).ConfigureAwait(false);
                return;
            }

            session = _sessionManager.GetSession(sessionId);
            if (session == null)
            {
                WriteJsonResponse(response, 404, "{\"error\":\"Session not found\"}");
                return;
            }

            wroteResponse = await session.HandlePostAsync(message, response.OutputStream, cancellationToken).ConfigureAwait(false);
            response.StatusCode = wroteResponse ? 200 : 202;
            response.Headers.Set("Mcp-Session-Id", sessionId);
        }

        private async Task HandleInitializePostAsync(JsonRpcMessage message, HttpListenerResponse response, CancellationToken cancellationToken)
        {
            JsonRpcRequest request;
            GridMcpSession session;
            string sessionId;
            bool wroteResponse;

            request = message as JsonRpcRequest;
            if (request == null || !string.Equals(request.Method, RequestMethods.Initialize, StringComparison.Ordinal))
            {
                WriteJsonResponse(response, 400, "{\"error\":\"The first MCP request must be initialize.\"}");
                return;
            }

            if (!_sessionManager.TryBeginInitialization())
            {
                WriteJsonResponse(response, 503, "{\"error\":\"An MCP session is already active.\"}");
                return;
            }

            session = null;
            try
            {
                sessionId = "grid-session-" + Guid.NewGuid().ToString("N");
                session = GridMcpSession.Create(_toolCatalog, sessionId);
                wroteResponse = await session.HandlePostAsync(message, response.OutputStream, cancellationToken).ConfigureAwait(false);
                response.StatusCode = wroteResponse ? 200 : 202;
                response.Headers.Set("Mcp-Session-Id", sessionId);
                _sessionManager.CompleteInitialization(session);
            }
            catch
            {
                _sessionManager.AbortInitialization();
                if (session != null)
                {
                    await session.DisposeAsync().ConfigureAwait(false);
                }

                throw;
            }
        }

        private async Task HandleGetAsync(HttpListenerRequest request, HttpListenerResponse response, CancellationToken cancellationToken)
        {
            string sessionId;
            GridMcpSession session;

            sessionId = request.Headers["Mcp-Session-Id"];
            session = _sessionManager.GetSession(sessionId);
            if (session == null)
            {
                WriteJsonResponse(response, 404, "{\"error\":\"Session not found\"}");
                return;
            }

            if ((request.Headers["Accept"] ?? string.Empty).IndexOf("text/event-stream", StringComparison.OrdinalIgnoreCase) < 0)
            {
                WriteJsonResponse(response, 400, "{\"error\":\"GET /mcp requires Accept: text/event-stream\"}");
                return;
            }

            response.StatusCode = 200;
            response.ContentType = "text/event-stream";
            response.Headers.Set("Cache-Control", "no-cache");
            response.Headers.Set("Connection", "keep-alive");
            response.SendChunked = true;
            response.Headers.Set("Mcp-Session-Id", sessionId);

            await session.HandleGetAsync(response.OutputStream, cancellationToken).ConfigureAwait(false);
        }

        private async Task HandleDeleteAsync(HttpListenerRequest request, HttpListenerResponse response)
        {
            string sessionId;
            GridMcpSession session;

            sessionId = request.Headers["Mcp-Session-Id"];
            session = _sessionManager.RemoveSession(sessionId);
            if (session != null)
            {
                await session.DisposeAsync().ConfigureAwait(false);
            }

            response.StatusCode = 200;
        }

        private static void WriteJsonResponse(HttpListenerResponse response, int statusCode, string json)
        {
            byte[] bytes;

            bytes = Encoding.UTF8.GetBytes(json);
            response.StatusCode = statusCode;
            response.ContentType = "application/json";
            response.ContentLength64 = bytes.Length;
            response.OutputStream.Write(bytes, 0, bytes.Length);
        }
    }
}
