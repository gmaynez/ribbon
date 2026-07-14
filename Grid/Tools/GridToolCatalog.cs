using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace Grid.Tools
{
    internal sealed class GridToolCatalog
    {
        private readonly Dictionary<string, GridToolDescriptor> _toolLookup;
        private readonly List<GridToolDescriptor> _tools;

        public GridToolCatalog(GridToolHandlers handlers)
        {
            BindingFlags bindingFlags;
            MethodInfo[] methods;
            int index;

            if (handlers == null)
            {
                throw new ArgumentNullException(nameof(handlers));
            }

            _toolLookup = new Dictionary<string, GridToolDescriptor>(StringComparer.Ordinal);
            _tools = new List<GridToolDescriptor>();

            bindingFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly;
            methods = typeof(GridToolHandlers).GetMethods(bindingFlags);

            for (index = 0; index < methods.Length; index++)
            {
                GridToolDescriptor descriptor;

                descriptor = GridToolDescriptor.Create(methods[index], handlers);
                if (descriptor == null)
                {
                    continue;
                }

                _tools.Add(descriptor);
                _toolLookup.Add(descriptor.Name, descriptor);
            }
        }

        public IEnumerable<GridToolDescriptor> Tools
        {
            get { return _tools; }
        }

        public McpServerPrimitiveCollection<McpServerTool> CreateMcpToolCollection()
        {
            McpServerPrimitiveCollection<McpServerTool> collection;
            int index;

            collection = new McpServerPrimitiveCollection<McpServerTool>(StringComparer.Ordinal);
            for (index = 0; index < _tools.Count; index++)
            {
                collection.Add(_tools[index].McpTool);
            }

            return collection;
        }

        public JsonArray CreateOpenAiToolDefinitions()
        {
            JsonArray tools;
            int index;

            tools = new JsonArray();
            for (index = 0; index < _tools.Count; index++)
            {
                tools.Add(_tools[index].CreateOpenAiToolDefinition());
            }

            return tools;
        }

        public async Task<ToolInvocationResult> InvokeAsync(string toolName, JsonObject arguments, CancellationToken cancellationToken)
        {
            GridToolDescriptor descriptor;
            object value;

            if (string.IsNullOrWhiteSpace(toolName))
            {
                throw new ArgumentException("Tool name is required.", nameof(toolName));
            }

            if (!_toolLookup.TryGetValue(toolName, out descriptor))
            {
                throw new McpException(string.Format("Unknown tool '{0}'.", toolName));
            }

            value = await descriptor.InvokeAsync(arguments ?? new JsonObject(), cancellationToken).ConfigureAwait(false);
            return ToolInvocationResult.FromSuccess(toolName, value);
        }

        internal sealed class GridToolDescriptor
        {
            private readonly MethodInfo _method;
            private readonly object _target;
            private readonly string _description;
            private readonly JsonNode _inputSchema;

            private GridToolDescriptor(MethodInfo method, object target, McpServerTool mcpTool, string description, JsonNode inputSchema)
            {
                _method = method;
                _target = target;
                McpTool = mcpTool;
                _description = description;
                _inputSchema = inputSchema;
            }

            public string Name
            {
                get { return McpTool.ProtocolTool.Name; }
            }

            public McpServerTool McpTool { get; private set; }

            public static GridToolDescriptor Create(MethodInfo method, object target)
            {
                McpServerToolAttribute attribute;
                McpServerTool mcpTool;
                object schemaObject;
                JsonNode schemaNode;
                string description;

                attribute = method.GetCustomAttribute<McpServerToolAttribute>();
                if (attribute == null)
                {
                    return null;
                }

                mcpTool = McpServerTool.Create(method, target, new McpServerToolCreateOptions
                {
                    SerializerOptions = McpJsonUtilities.DefaultOptions
                });

                schemaObject = mcpTool.ProtocolTool.InputSchema;
                schemaNode = schemaObject == null
                    ? new JsonObject { ["type"] = "object" }
                    : JsonNode.Parse(JsonSerializer.Serialize(schemaObject, McpJsonUtilities.DefaultOptions));

                description = mcpTool.ProtocolTool.Description;
                if (string.IsNullOrWhiteSpace(description))
                {
                    DescriptionAttribute descriptionAttribute;
                    descriptionAttribute = method.GetCustomAttribute<DescriptionAttribute>();
                    description = descriptionAttribute != null ? descriptionAttribute.Description : method.Name;
                }

                return new GridToolDescriptor(method, target, mcpTool, description, schemaNode);
            }

            public JsonObject CreateOpenAiToolDefinition()
            {
                return new JsonObject
                {
                    ["type"] = "function",
                    ["function"] = new JsonObject
                    {
                        ["name"] = Name,
                        ["description"] = _description,
                        ["parameters"] = _inputSchema == null ? new JsonObject { ["type"] = "object" } : _inputSchema.DeepClone()
                    }
                };
            }

            public async Task<object> InvokeAsync(JsonObject arguments, CancellationToken cancellationToken)
            {
                ParameterInfo[] parameters;
                object[] values;
                object invocationResult;
                Task task;
                PropertyInfo resultProperty;
                int index;

                parameters = _method.GetParameters();
                values = new object[parameters.Length];

                for (index = 0; index < parameters.Length; index++)
                {
                    values[index] = BindParameter(parameters[index], arguments, cancellationToken);
                }

                try
                {
                    invocationResult = _method.Invoke(_target, values);
                }
                catch (TargetInvocationException ex)
                {
                    if (ex.InnerException != null)
                    {
                        throw ex.InnerException;
                    }

                    throw;
                }

                task = invocationResult as Task;
                if (task == null)
                {
                    return invocationResult;
                }

                await task.ConfigureAwait(false);
                resultProperty = task.GetType().GetProperty("Result");
                return resultProperty != null ? resultProperty.GetValue(task) : null;
            }

            private static object BindParameter(ParameterInfo parameter, JsonObject arguments, CancellationToken cancellationToken)
            {
                JsonNode node;
                string rawJson;

                if (parameter.ParameterType == typeof(CancellationToken))
                {
                    return cancellationToken;
                }

                if (arguments != null && arguments.TryGetPropertyValue(parameter.Name, out node) && node != null)
                {
                    rawJson = node.ToJsonString();
                    return JsonSerializer.Deserialize(rawJson, parameter.ParameterType, McpJsonUtilities.DefaultOptions);
                }

                if (parameter.HasDefaultValue)
                {
                    return parameter.DefaultValue;
                }

                throw new McpException(string.Format("Parameter '{0}' is required.", parameter.Name));
            }
        }
    }

    internal sealed class ToolInvocationResult
    {
        private ToolInvocationResult(string toolName, object data, bool success, string content)
        {
            ToolName = toolName;
            Data = data;
            Success = success;
            Content = content;
        }

        public string ToolName { get; private set; }

        public object Data { get; private set; }

        public bool Success { get; private set; }

        public string Content { get; private set; }

        public static ToolInvocationResult FromSuccess(string toolName, object data)
        {
            return new ToolInvocationResult(toolName, data, true, JsonSerializer.Serialize(data, McpJsonUtilities.DefaultOptions));
        }

        public static ToolInvocationResult FromError(string toolName, string message)
        {
            Dictionary<string, object> payload;

            payload = new Dictionary<string, object>
            {
                ["success"] = false,
                ["error"] = message
            };

            return new ToolInvocationResult(toolName, payload, false, JsonSerializer.Serialize(payload, McpJsonUtilities.DefaultOptions));
        }
    }
}
