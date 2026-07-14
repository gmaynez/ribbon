using System;
using System.IO;
using System.Net;
using System.Text;

namespace Grid.Mcp
{
    internal sealed class LocalRequestValidator
    {
        public bool Validate(HttpListenerRequest request, HttpListenerResponse response)
        {
            return ValidateHost(request, response) && ValidateOrigin(request, response);
        }

        private static bool ValidateHost(HttpListenerRequest request, HttpListenerResponse response)
        {
            string host;
            string hostname;
            int separatorIndex;

            host = request.UserHostName;
            if (string.IsNullOrWhiteSpace(host))
            {
                WriteJsonError(response, 400, "Missing Host header.");
                return false;
            }

            if (host.StartsWith("[", StringComparison.Ordinal))
            {
                separatorIndex = host.IndexOf(']');
                hostname = separatorIndex >= 0 ? host.Substring(0, separatorIndex + 1) : host;
            }
            else
            {
                separatorIndex = host.IndexOf(':');
                hostname = separatorIndex >= 0 ? host.Substring(0, separatorIndex) : host;
            }

            if (!IsLoopbackHost(hostname))
            {
                WriteJsonError(response, 403, "Only localhost requests are allowed.");
                return false;
            }

            return true;
        }

        private static bool ValidateOrigin(HttpListenerRequest request, HttpListenerResponse response)
        {
            string origin;
            Uri originUri;

            origin = request.Headers["Origin"];
            if (string.IsNullOrWhiteSpace(origin))
            {
                return true;
            }

            if (!Uri.TryCreate(origin, UriKind.Absolute, out originUri) || !IsLoopbackHost(originUri.Host))
            {
                WriteJsonError(response, 403, "Only localhost origins are allowed.");
                return false;
            }

            return true;
        }

        private static bool IsLoopbackHost(string hostname)
        {
            return string.Equals(hostname, "localhost", StringComparison.OrdinalIgnoreCase)
                || string.Equals(hostname, "127.0.0.1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(hostname, "[::1]", StringComparison.OrdinalIgnoreCase)
                || string.Equals(hostname, "::1", StringComparison.OrdinalIgnoreCase);
        }

        private static void WriteJsonError(HttpListenerResponse response, int statusCode, string message)
        {
            byte[] bytes;
            string payload;

            payload = "{\"error\":\"" + Escape(message) + "\"}";
            bytes = Encoding.UTF8.GetBytes(payload);

            response.StatusCode = statusCode;
            response.ContentType = "application/json";
            response.ContentLength64 = bytes.Length;

            try
            {
                response.OutputStream.Write(bytes, 0, bytes.Length);
            }
            catch (IOException)
            {
            }
        }

        private static string Escape(string value)
        {
            return (value ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"");
        }
    }
}
