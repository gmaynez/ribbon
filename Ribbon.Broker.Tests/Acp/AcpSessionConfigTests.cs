using System.Text.Json;
using Ribbon.Broker.Acp;
using Xunit;

namespace Ribbon.Broker.Tests.Acp;

public sealed class AcpSessionConfigTests
{
    [Fact]
    public void BrokerPipeNameTracksTheWireProtocolVersion()
    {
        Assert.EndsWith(".v" + Ribbon.Contracts.RibbonProtocol.Version, Ribbon.Contracts.RibbonProtocol.PipeName);
        Assert.EndsWith(".v" + Ribbon.Contracts.RibbonProtocol.Version, Ribbon.Contracts.RibbonProtocol.BrokerMutexName);
    }

    [Fact]
    public void ParsePreservesAgentModelNamesValuesAndOrder()
    {
        using var document = JsonDocument.Parse("""
            {
              "configOptions": [
                {
                  "id": "model",
                  "name": "Model",
                  "category": "model",
                  "type": "select",
                  "currentValue": "fast",
                  "options": [
                    { "value": "fast", "name": "Fast", "description": "Lower latency" },
                    { "value": "deep", "name": "Deep", "description": "More reasoning" }
                  ]
                }
              ]
            }
            """);

        var options = AcpSessionConfig.Parse(document.RootElement);

        var model = Assert.Single(options);
        Assert.Equal("model", model.Id);
        Assert.Equal("model", model.Category);
        Assert.Equal("fast", model.CurrentValue);
        Assert.Equal(["fast", "deep"], model.Options.Select(value => value.Value));
        Assert.Equal(["Fast", "Deep"], model.Options.Select(value => value.Name));
    }

    [Fact]
    public void ParseIgnoresMalformedOptionsWithoutRejectingTheSession()
    {
        using var document = JsonDocument.Parse("""
            {
              "configOptions": [
                { "id": "missing-name", "type": "select", "currentValue": "x", "options": [] },
                { "id": "model", "name": "Model", "type": "select", "currentValue": "x", "options": [
                  { "value": "x", "name": "Model X" }
                ] }
              ]
            }
            """);

        var model = Assert.Single(AcpSessionConfig.Parse(document.RootElement));

        Assert.Equal("model", model.Id);
    }

    [Fact]
    public void RequireSelectValueRejectsValuesTheAgentDidNotAdvertise()
    {
        var options = new[]
        {
            new Ribbon.Contracts.SessionConfigOption
            {
                Id = "model",
                Type = "select",
                Options = [new Ribbon.Contracts.SessionConfigOptionValue { Value = "available", Name = "Available" }]
            }
        };

        Assert.Throws<ArgumentOutOfRangeException>(() => AcpSessionConfig.RequireSelectValue(options, "model", "invented"));
    }
}
