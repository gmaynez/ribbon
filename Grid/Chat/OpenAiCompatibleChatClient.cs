using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using Grid.Constants;

namespace Grid.Chat
{
    internal sealed class OpenAiCompatibleChatClient : IDisposable
    {
        private readonly HttpClient _httpClient;
        private bool _disposed;

        public OpenAiCompatibleChatClient()
        {
            _httpClient = new HttpClient();
            _httpClient.Timeout = GridConstants.ProviderTimeout;
        }

        public async Task<OpenAiChatResponse> CreateChatCompletionAsync(
            string baseUrl,
            string apiKey,
            string model,
            IEnumerable<JsonObject> messages,
            JsonArray tools,
            CancellationToken cancellationToken)
        {
            HttpRequestMessage request;
            HttpResponseMessage response;
            JsonArray messageArray;
            JsonObject payload;
            JsonObject responseObject;
            JsonObject choice;
            JsonObject message;
            string responseText;

            if (string.IsNullOrWhiteSpace(baseUrl))
            {
                throw new InvalidOperationException("Provider base URL is required.");
            }

            if (string.IsNullOrWhiteSpace(apiKey))
            {
                throw new InvalidOperationException("Provider API key is required.");
            }

            if (string.IsNullOrWhiteSpace(model))
            {
                throw new InvalidOperationException("Provider model is required.");
            }

            messageArray = new JsonArray();
            foreach (JsonObject messageObject in messages)
            {
                messageArray.Add(CloneNode(messageObject));
            }

            payload = new JsonObject
            {
                ["model"] = model,
                ["messages"] = messageArray,
                ["tool_choice"] = "auto",
                ["temperature"] = 0.2,
                ["stream"] = false,
                ["tools"] = CloneNode(tools)
            };

            request = new HttpRequestMessage(HttpMethod.Post, BuildChatCompletionsUri(baseUrl));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Content = new StringContent(payload.ToJsonString(), Encoding.UTF8, "application/json");

            response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
            responseText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException(ExtractProviderError(responseText, (int)response.StatusCode));
            }

            responseObject = JsonNode.Parse(responseText) as JsonObject;
            if (responseObject == null)
            {
                throw new InvalidOperationException("Provider response did not contain a valid JSON object.");
            }

            choice = responseObject["choices"]?[0] as JsonObject;
            message = choice != null ? choice["message"] as JsonObject : null;
            if (message == null)
            {
                throw new InvalidOperationException("Provider response did not contain a chat message.");
            }

            return new OpenAiChatResponse(ExtractContentText(message["content"]), ParseToolCalls(message["tool_calls"] as JsonArray));
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _httpClient.Dispose();
        }

        private static Uri BuildChatCompletionsUri(string baseUrl)
        {
            string normalized;

            normalized = baseUrl.Trim();
            if (normalized.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
            {
                return new Uri(normalized, UriKind.Absolute);
            }

            if (!normalized.EndsWith("/", StringComparison.Ordinal))
            {
                normalized += "/";
            }

            return new Uri(new Uri(normalized, UriKind.Absolute), "chat/completions");
        }

        private static string ExtractProviderError(string responseText, int statusCode)
        {
            JsonObject responseObject;
            JsonObject errorObject;
            string message;

            responseObject = JsonNode.Parse(responseText) as JsonObject;
            errorObject = responseObject != null ? responseObject["error"] as JsonObject : null;
            message = errorObject != null ? errorObject["message"] != null ? errorObject["message"].ToString() : null : null;

            if (string.IsNullOrWhiteSpace(message))
            {
                message = responseText;
            }

            return string.Format("Provider request failed ({0}): {1}", statusCode, message);
        }

        private static string ExtractContentText(JsonNode contentNode)
        {
            JsonArray contentArray;
            StringBuilder builder;

            if (contentNode == null)
            {
                return string.Empty;
            }

            if (contentNode is JsonValue)
            {
                return contentNode.ToString();
            }

            contentArray = contentNode as JsonArray;
            if (contentArray == null)
            {
                return contentNode.ToJsonString();
            }

            builder = new StringBuilder();
            foreach (JsonNode node in contentArray)
            {
                JsonObject contentObject;
                string type;

                contentObject = node as JsonObject;
                if (contentObject == null)
                {
                    continue;
                }

                type = contentObject["type"] != null ? contentObject["type"].ToString() : string.Empty;
                if (string.Equals(type, "text", StringComparison.OrdinalIgnoreCase) && contentObject["text"] != null)
                {
                    if (builder.Length > 0)
                    {
                        builder.AppendLine();
                    }

                    builder.Append(contentObject["text"].ToString());
                }
            }

            return builder.ToString();
        }

        private static List<OpenAiToolCall> ParseToolCalls(JsonArray toolCallsArray)
        {
            List<OpenAiToolCall> toolCalls;

            toolCalls = new List<OpenAiToolCall>();
            if (toolCallsArray == null)
            {
                return toolCalls;
            }

            foreach (JsonNode node in toolCallsArray)
            {
                JsonObject toolCallObject;
                JsonObject functionObject;
                string id;
                string name;
                string argumentsJson;

                toolCallObject = node as JsonObject;
                functionObject = toolCallObject != null ? toolCallObject["function"] as JsonObject : null;
                if (toolCallObject == null || functionObject == null)
                {
                    continue;
                }

                id = toolCallObject["id"] != null ? toolCallObject["id"].ToString() : Guid.NewGuid().ToString("N");
                name = functionObject["name"] != null ? functionObject["name"].ToString() : string.Empty;
                argumentsJson = functionObject["arguments"] != null ? functionObject["arguments"].ToString() : "{}";

                toolCalls.Add(new OpenAiToolCall(id, name, argumentsJson));
            }

            return toolCalls;
        }

        private static JsonNode CloneNode(JsonNode node)
        {
            return node == null ? null : JsonNode.Parse(node.ToJsonString());
        }
    }

    internal sealed class OpenAiChatResponse
    {
        public OpenAiChatResponse(string content, IList<OpenAiToolCall> toolCalls)
        {
            Content = content ?? string.Empty;
            ToolCalls = toolCalls ?? new List<OpenAiToolCall>();
        }

        public string Content { get; private set; }

        public IList<OpenAiToolCall> ToolCalls { get; private set; }
    }

    internal sealed class OpenAiToolCall
    {
        public OpenAiToolCall(string id, string name, string argumentsJson)
        {
            Id = id;
            Name = name;
            ArgumentsJson = argumentsJson;
        }

        public string Id { get; private set; }

        public string Name { get; private set; }

        public string ArgumentsJson { get; private set; }
    }
}
