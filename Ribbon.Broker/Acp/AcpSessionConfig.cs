using System.Text.Json;
using Ribbon.Contracts;

namespace Ribbon.Broker.Acp;

internal static class AcpSessionConfig
{
    public static List<SessionConfigOption> Parse(JsonElement container)
    {
        if (!container.TryGetProperty("configOptions", out var configOptions)
            || configOptions.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var result = new List<SessionConfigOption>();
        foreach (var option in configOptions.EnumerateArray())
        {
            if (!TryGetRequiredString(option, "id", out var id)
                || !TryGetRequiredString(option, "name", out var name)
                || !TryGetRequiredString(option, "type", out var type))
            {
                continue;
            }

            var currentValue = option.TryGetProperty("currentValue", out var current)
                && current.ValueKind == JsonValueKind.String
                    ? current.GetString() ?? string.Empty
                    : string.Empty;
            var values = new List<SessionConfigOptionValue>();
            if (option.TryGetProperty("options", out var options) && options.ValueKind == JsonValueKind.Array)
            {
                foreach (var value in options.EnumerateArray())
                {
                    if (!TryGetRequiredString(value, "value", out var valueId)
                        || !TryGetRequiredString(value, "name", out var valueName))
                    {
                        continue;
                    }
                    values.Add(new SessionConfigOptionValue
                    {
                        Value = valueId,
                        Name = valueName,
                        Description = OptionalString(value, "description")
                    });
                }
            }

            result.Add(new SessionConfigOption
            {
                Id = id,
                Name = name,
                Description = OptionalString(option, "description"),
                Category = OptionalString(option, "category"),
                Type = type,
                CurrentValue = currentValue,
                Options = values
            });
        }
        return result;
    }

    public static SessionConfigOption RequireSelectValue(
        IReadOnlyList<SessionConfigOption> configOptions,
        string configId,
        string value)
    {
        if (string.IsNullOrWhiteSpace(configId) || string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException("A configuration option and value are required.");
        }

        var option = configOptions.FirstOrDefault(item => string.Equals(item.Id, configId, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"The ACP session does not expose configuration option '{configId}'.");
        if (!string.Equals(option.Type, "select", StringComparison.Ordinal))
        {
            throw new NotSupportedException($"Ribbon cannot change ACP configuration option type '{option.Type}'.");
        }
        if (option.Options == null || !option.Options.Any(item => string.Equals(item.Value, value, StringComparison.Ordinal)))
        {
            throw new ArgumentOutOfRangeException(nameof(value), $"'{value}' is not available for ACP configuration option '{configId}'.");
        }
        return option;
    }

    private static bool TryGetRequiredString(JsonElement element, string propertyName, out string value)
    {
        value = string.Empty;
        if (!element.TryGetProperty(propertyName, out var property) || property.ValueKind != JsonValueKind.String)
        {
            return false;
        }
        value = property.GetString() ?? string.Empty;
        return !string.IsNullOrWhiteSpace(value);
    }

    private static string OptionalString(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.String
            ? property.GetString() ?? string.Empty
            : string.Empty;
    }
}
