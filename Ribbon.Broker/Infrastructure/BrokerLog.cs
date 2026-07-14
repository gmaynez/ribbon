namespace Ribbon.Broker.Infrastructure;

internal sealed class BrokerLog
{
    private readonly string _path;
    private readonly object _gate = new();

    public BrokerLog(BrokerPaths paths)
    {
        _path = Path.Combine(paths.Logs, $"broker-{DateTime.UtcNow:yyyyMMdd}.log");
    }

    public void Info(string message)
    {
        Write("INFO", message);
    }

    public void Error(string message, Exception? exception = null)
    {
        Write("ERROR", exception == null ? message : message + Environment.NewLine + exception);
    }

    private void Write(string level, string message)
    {
        var line = $"{DateTimeOffset.UtcNow:O} [{level}] {message}{Environment.NewLine}";
        lock (_gate)
        {
            File.AppendAllText(_path, line);
        }
    }
}
