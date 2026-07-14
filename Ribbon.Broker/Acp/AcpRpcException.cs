namespace Ribbon.Broker.Acp;

internal sealed class AcpRpcException : Exception
{
    public AcpRpcException(int code, string message, string? data)
        : base(message)
    {
        Code = code;
        DataJson = data;
    }

    public int Code { get; }
    public string? DataJson { get; }
}
