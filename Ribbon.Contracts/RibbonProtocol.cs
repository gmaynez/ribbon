using System;
using System.Collections.Generic;

namespace Ribbon.Contracts
{
    public static class RibbonProtocol
    {
        public const int Version = 1;
        public const string PipeName = "Ribbon.Broker.v1";
        public const string ProductVersion = "0.1.0";

        public const string RegisterHost = "host/register";
        public const string UnregisterHost = "host/unregister";
        public const string ListTools = "host/tools/list";
        public const string InvokeTool = "host/tools/call";
        public const string ListInstalledAgents = "agents/list";
        public const string ListRegistryAgents = "agents/registry";
        public const string InstallAgent = "agents/install";
        public const string UninstallAgent = "agents/uninstall";
        public const string AuthenticateAgent = "agents/authenticate";
        public const string StartSession = "session/start";
        public const string PromptSession = "session/prompt";
        public const string CancelSession = "session/cancel";
        public const string SessionUpdate = "session/update";
        public const string PermissionRequest = "session/request_permission";
    }

    public sealed class RpcEnvelope
    {
        public int Version { get; set; }
        public string Kind { get; set; }
        public string Id { get; set; }
        public string Method { get; set; }
        public string Payload { get; set; }
        public bool Success { get; set; }
        public string Error { get; set; }

        public static RpcEnvelope Request(string method, string payload)
        {
            return new RpcEnvelope
            {
                Version = RibbonProtocol.Version,
                Kind = "request",
                Id = Guid.NewGuid().ToString("N"),
                Method = method,
                Payload = payload ?? string.Empty,
                Success = true
            };
        }

        public static RpcEnvelope Notification(string method, string payload)
        {
            return new RpcEnvelope
            {
                Version = RibbonProtocol.Version,
                Kind = "notification",
                Id = string.Empty,
                Method = method,
                Payload = payload ?? string.Empty,
                Success = true
            };
        }

        public static RpcEnvelope Response(RpcEnvelope request, string payload)
        {
            return new RpcEnvelope
            {
                Version = RibbonProtocol.Version,
                Kind = "response",
                Id = request != null ? request.Id : string.Empty,
                Method = request != null ? request.Method : string.Empty,
                Payload = payload ?? string.Empty,
                Success = true
            };
        }

        public static RpcEnvelope Failure(RpcEnvelope request, string error)
        {
            return new RpcEnvelope
            {
                Version = RibbonProtocol.Version,
                Kind = "response",
                Id = request != null ? request.Id : string.Empty,
                Method = request != null ? request.Method : string.Empty,
                Payload = string.Empty,
                Success = false,
                Error = error ?? "Unknown broker error."
            };
        }
    }

    public sealed class HostRegistration
    {
        public string HostId { get; set; }
        public string HostKind { get; set; }
        public int ProcessId { get; set; }
        public string DisplayName { get; set; }
        public string DocumentPath { get; set; }
        public string Version { get; set; }
    }

    public sealed class OfficeToolDefinition
    {
        public string HostId { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string InputSchemaJson { get; set; }
        public bool Destructive { get; set; }
        public string HostKind { get; set; }
    }

    public sealed class OfficeToolInvocation
    {
        public string HostId { get; set; }
        public string ToolName { get; set; }
        public string ArgumentsJson { get; set; }
    }

    public sealed class OfficeToolResult
    {
        public bool Success { get; set; }
        public string ContentJson { get; set; }
        public string Error { get; set; }
    }

    public sealed class AgentSummary
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Version { get; set; }
        public string Description { get; set; }
        public string Command { get; set; }
        public IList<string> Arguments { get; set; }
        public bool Installed { get; set; }
        public bool UpdateAvailable { get; set; }
        public string License { get; set; }
        public string Website { get; set; }
        public string DistributionType { get; set; }
    }

    public sealed class AgentIdRequest
    {
        public string AgentId { get; set; }
    }

    public sealed class HostIdRequest
    {
        public string HostId { get; set; }
    }

    public sealed class AgentAuthenticationRequest
    {
        public string AgentId { get; set; }
        public string MethodId { get; set; }
    }

    public sealed class AgentAuthenticationMethod
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
    }

    public sealed class SessionStartRequest
    {
        public string AgentId { get; set; }
        public string HostId { get; set; }
        public string WorkingDirectory { get; set; }
    }

    public sealed class SessionStartResponse
    {
        public string SessionId { get; set; }
        public string AgentName { get; set; }
        public IList<AgentAuthenticationMethod> AuthenticationMethods { get; set; }
    }

    public sealed class SessionPromptRequest
    {
        public string SessionId { get; set; }
        public string Text { get; set; }
    }

    public sealed class SessionCancelRequest
    {
        public string SessionId { get; set; }
    }

    public sealed class SessionUpdateMessage
    {
        public string SessionId { get; set; }
        public string UpdateKind { get; set; }
        public string Text { get; set; }
        public string ToolName { get; set; }
        public string Status { get; set; }
        public string RawJson { get; set; }
    }

    public sealed class PermissionPrompt
    {
        public string SessionId { get; set; }
        public string ToolCallId { get; set; }
        public string Title { get; set; }
        public string RawJson { get; set; }
        public IList<PermissionChoice> Options { get; set; }
    }

    public sealed class PermissionChoice
    {
        public string OptionId { get; set; }
        public string Name { get; set; }
        public string Kind { get; set; }
    }

    public sealed class PermissionDecision
    {
        public string OptionId { get; set; }
        public bool Cancelled { get; set; }
    }
}
