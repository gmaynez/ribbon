using Ribbon.Broker.Acp;
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
}
