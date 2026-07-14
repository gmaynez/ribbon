using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using Grid.Configuration;
using Grid.Constants;
using Grid.Tools;

namespace Grid.Chat
{
    internal sealed class GridConversationService : IDisposable
    {
        private readonly OpenAiCompatibleChatClient _chatClient;
        private readonly GridSettings _settings;
        private readonly GridToolCatalog _toolCatalog;
        private readonly List<JsonObject> _messages;
        private bool _disposed;

        public GridConversationService(GridSettings settings, GridToolCatalog toolCatalog)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _toolCatalog = toolCatalog ?? throw new ArgumentNullException(nameof(toolCatalog));
            _chatClient = new OpenAiCompatibleChatClient();
            _messages = new List<JsonObject>();
        }

        public async Task<ConversationTurnResult> SendAsync(string userMessage, CancellationToken cancellationToken)
        {
            JsonArray toolDefinitions;
            List<string> executedTools;
            int round;

            if (string.IsNullOrWhiteSpace(userMessage))
            {
                throw new InvalidOperationException("A message is required.");
            }

            EnsureSystemMessage();
            _messages.Add(new JsonObject
            {
                ["role"] = "user",
                ["content"] = userMessage.Trim()
            });

            toolDefinitions = _toolCatalog.CreateOpenAiToolDefinitions();
            executedTools = new List<string>();

            for (round = 0; round < GridConstants.MaxAgentToolRounds; round++)
            {
                OpenAiChatResponse response;
                JsonObject assistantMessage;

                response = await _chatClient.CreateChatCompletionAsync(
                    _settings.ProviderBaseUrl,
                    _settings.GetProviderApiKey(),
                    _settings.ProviderModel,
                    _messages,
                    toolDefinitions,
                    cancellationToken).ConfigureAwait(false);

                assistantMessage = new JsonObject
                {
                    ["role"] = "assistant"
                };

                if (!string.IsNullOrWhiteSpace(response.Content))
                {
                    assistantMessage["content"] = response.Content;
                }

                if (response.ToolCalls.Count > 0)
                {
                    assistantMessage["tool_calls"] = CreateToolCallArray(response.ToolCalls);
                }

                _messages.Add(assistantMessage);

                if (response.ToolCalls.Count == 0)
                {
                    return new ConversationTurnResult(response.Content, executedTools);
                }

                foreach (OpenAiToolCall toolCall in response.ToolCalls)
                {
                    ToolInvocationResult invocationResult;
                    JsonObject arguments;

                    executedTools.Add(toolCall.Name);
                    arguments = ParseArguments(toolCall.ArgumentsJson);

                    try
                    {
                        invocationResult = await _toolCatalog.InvokeAsync(toolCall.Name, arguments, cancellationToken).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        invocationResult = ToolInvocationResult.FromError(toolCall.Name, ex.Message);
                    }

                    _messages.Add(new JsonObject
                    {
                        ["role"] = "tool",
                        ["tool_call_id"] = toolCall.Id,
                        ["name"] = toolCall.Name,
                        ["content"] = invocationResult.Content
                    });
                }
            }

            throw new InvalidOperationException("The assistant exceeded the maximum number of tool rounds for a single response.");
        }

        public void ClearConversation()
        {
            _messages.Clear();
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _chatClient.Dispose();
        }

        private void EnsureSystemMessage()
        {
            if (_messages.Count > 0)
            {
                return;
            }

            _messages.Add(new JsonObject
            {
                ["role"] = "system",
                ["content"] = _settings.SystemPrompt
            });
        }

        private static JsonObject ParseArguments(string argumentsJson)
        {
            JsonObject arguments;

            if (string.IsNullOrWhiteSpace(argumentsJson))
            {
                return new JsonObject();
            }

            arguments = JsonNode.Parse(argumentsJson) as JsonObject;
            return arguments ?? new JsonObject();
        }

        private static JsonArray CreateToolCallArray(IEnumerable<OpenAiToolCall> toolCalls)
        {
            JsonArray array;

            array = new JsonArray();
            foreach (OpenAiToolCall toolCall in toolCalls)
            {
                array.Add(new JsonObject
                {
                    ["id"] = toolCall.Id,
                    ["type"] = "function",
                    ["function"] = new JsonObject
                    {
                        ["name"] = toolCall.Name,
                        ["arguments"] = toolCall.ArgumentsJson
                    }
                });
            }

            return array;
        }
    }

    internal sealed class ConversationTurnResult
    {
        public ConversationTurnResult(string assistantMessage, IList<string> executedTools)
        {
            AssistantMessage = assistantMessage ?? string.Empty;
            ExecutedTools = executedTools != null ? executedTools.ToList() : new List<string>();
        }

        public string AssistantMessage { get; private set; }

        public IList<string> ExecutedTools { get; private set; }
    }
}
