using System.Text.Json;

namespace Ribbon.Broker.Infrastructure;

internal static class JsonCodec
{
    public static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = false
    };

    public static string Serialize<T>(T value)
    {
        return JsonSerializer.Serialize(value, Options);
    }

    public static T Deserialize<T>(string? json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            throw new InvalidDataException("A JSON payload was required.");
        }

        return JsonSerializer.Deserialize<T>(json, Options)
            ?? throw new InvalidDataException($"Unable to deserialize {typeof(T).Name}.");
    }
}
