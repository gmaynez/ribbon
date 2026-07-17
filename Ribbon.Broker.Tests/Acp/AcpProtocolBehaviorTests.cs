using Ribbon.Broker.Acp;
using System.Text.Json;
using Xunit;

namespace Ribbon.Broker.Tests.Acp;

public sealed class AcpProtocolBehaviorTests
{
    [Fact]
    public void AuthenticationIsOfferedOnlyForTheAcpAuthenticationRequiredError()
    {
        Assert.True(AgentSessionManager.IsAuthenticationRequired(new AcpRpcException(-32000, "Authentication required.", null), 1));
        Assert.False(AgentSessionManager.IsAuthenticationRequired(new AcpRpcException(-32603, "Agent failed.", null), 1));
        Assert.False(AgentSessionManager.IsAuthenticationRequired(new AcpRpcException(-32000, "Authentication required.", null), 0));
    }

    [Fact]
    public void AcpRpcExceptionsPreserveTheirJsonRpcErrorCode()
    {
        var error = AcpProcessConnection.MapError(new AcpRpcException(-32601, "Method not found.", null));

        Assert.Equal(-32601, error.Code);
        Assert.Equal("Method not found.", error.Message);
    }

    [Fact]
    public void UnexpectedExceptionsBecomeInternalErrors()
    {
        var error = AcpProcessConnection.MapError(new InvalidOperationException("Unexpected failure."));

        Assert.Equal(-32603, error.Code);
        Assert.Equal("Unexpected failure.", error.Message);
    }

    [Fact]
    public void CancellingPermissionsAffectsPendingRequestsButNotFutureRequests()
    {
        var registry = new PendingPermissionRegistry();
        using var pending = registry.Register(CancellationToken.None);

        registry.CancelAll();

        Assert.True(pending.CancelledByClient);
        Assert.True(pending.Token.IsCancellationRequested);

        using var future = registry.Register(CancellationToken.None);
        Assert.False(future.CancelledByClient);
        Assert.False(future.Token.IsCancellationRequested);
    }

    [Fact]
    public void ThoughtChunksPreserveMessageIdentityAndText()
    {
        using var document = JsonDocument.Parse("""
            {
              "sessionUpdate": "agent_thought_chunk",
              "messageId": "thought-1",
              "content": { "type": "text", "text": "Inspecting the current document…" }
            }
            """);

        var update = AgentSessionManager.ParseSessionUpdate("session-1", document.RootElement);

        Assert.Equal("agent_thought_chunk", update.UpdateKind);
        Assert.Equal("thought-1", update.MessageId);
        Assert.Equal("Inspecting the current document…", update.Text);
    }

    [Fact]
    public void ToolUpdatesExposeProgressContentAndStatus()
    {
        using var document = JsonDocument.Parse("""
            {
              "sessionUpdate": "tool_call_update",
              "toolCallId": "call-1",
              "title": "Format the sales table",
              "kind": "edit",
              "status": "completed",
              "content": [
                {
                  "type": "content",
                  "content": { "type": "text", "text": "Applied the requested table style." }
                }
              ]
            }
            """);

        var update = AgentSessionManager.ParseSessionUpdate("session-1", document.RootElement);

        Assert.Equal("call-1", update.ToolCallId);
        Assert.Equal("Format the sales table", update.ToolName);
        Assert.Equal("edit", update.ToolKind);
        Assert.Equal("completed", update.Status);
        Assert.Equal("Applied the requested table style.", update.Text);
    }

    [Fact]
    public void PlansExposeEntriesForInlineProgressRendering()
    {
        using var document = JsonDocument.Parse("""
            {
              "sessionUpdate": "plan",
              "entries": [
                { "content": "Inspect the workbook", "priority": "high", "status": "completed" },
                { "content": "Create the chart", "priority": "medium", "status": "in_progress" }
              ]
            }
            """);

        var update = AgentSessionManager.ParseSessionUpdate("session-1", document.RootElement);

        Assert.Collection(
            update.PlanEntries!,
            entry =>
            {
                Assert.Equal("Inspect the workbook", entry.Content);
                Assert.Equal("completed", entry.Status);
            },
            entry =>
            {
                Assert.Equal("Create the chart", entry.Content);
                Assert.Equal("in_progress", entry.Status);
            });
    }

    [Fact]
    public void SessionCloseIsUsedOnlyWhenAdvertisedByTheAgent()
    {
        using var supported = JsonDocument.Parse("""
            { "agentCapabilities": { "sessionCapabilities": { "close": {} } } }
            """);
        using var unsupported = JsonDocument.Parse("""
            { "agentCapabilities": { "sessionCapabilities": {} } }
            """);

        Assert.True(AgentRuntime.SupportsSessionCloseCapability(supported.RootElement));
        Assert.False(AgentRuntime.SupportsSessionCloseCapability(unsupported.RootElement));
    }

    [Fact]
    public void ConversationHistoryCapabilitiesAreParsedIndependently()
    {
        using var supported = JsonDocument.Parse("""
            {
              "agentCapabilities": {
                "loadSession": true,
                "sessionCapabilities": {
                  "resume": {},
                  "list": {}
                }
              }
            }
            """);
        using var unsupported = JsonDocument.Parse("""
            { "agentCapabilities": { "loadSession": false, "sessionCapabilities": {} } }
            """);

        Assert.True(AgentRuntime.SupportsSessionLoadCapability(supported.RootElement));
        Assert.True(AgentRuntime.SupportsSessionCapability(supported.RootElement, "resume"));
        Assert.True(AgentRuntime.SupportsSessionCapability(supported.RootElement, "list"));
        Assert.False(AgentRuntime.SupportsSessionLoadCapability(unsupported.RootElement));
        Assert.False(AgentRuntime.SupportsSessionCapability(unsupported.RootElement, "resume"));
        Assert.False(AgentRuntime.SupportsSessionCapability(unsupported.RootElement, "list"));
    }

    [Fact]
    public void SessionInfoUpdatesExposeGeneratedConversationMetadata()
    {
        using var document = JsonDocument.Parse("""
            {
              "sessionUpdate": "session_info_update",
              "title": "Build the quarterly forecast",
              "updatedAt": "2026-07-16T18:42:00Z"
            }
            """);

        var update = AgentSessionManager.ParseSessionUpdate("session-1", document.RootElement);

        Assert.Equal("session_info_update", update.UpdateKind);
        Assert.Equal("Build the quarterly forecast", update.Title);
        Assert.Equal("2026-07-16T18:42:00Z", update.UpdatedAt);
    }
}
