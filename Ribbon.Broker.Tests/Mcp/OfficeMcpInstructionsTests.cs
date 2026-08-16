using Ribbon.Broker.Mcp;
using Ribbon.Contracts;
using Xunit;

namespace Ribbon.Broker.Tests.Mcp;

public sealed class OfficeMcpInstructionsTests
{
    [Fact]
    public void EmptyCatalogExplainsThatNoOfficeToolsAreAvailable()
    {
        var instructions = OfficeMcpInstructions.Build([]);

        Assert.Contains("No Office tools are currently available", instructions);
        Assert.DoesNotContain("Excel workflow", instructions);
        Assert.DoesNotContain("Word workflow", instructions);
        Assert.DoesNotContain("PowerPoint workflow", instructions);
    }

    [Fact]
    public void ExcelOnlyCatalogContainsOnlyAvailableExcelGuidance()
    {
        var instructions = OfficeMcpInstructions.Build(Tools("excel_get_context", "excel_read_range", "excel_format_range"));

        Assert.Contains("Excel workflow", instructions);
        Assert.Contains("excel_get_context", instructions);
        Assert.Contains("excel_read_range", instructions);
        Assert.Contains("excel_format_range", instructions);
        Assert.DoesNotContain("excel_write_formulas", instructions);
        Assert.DoesNotContain("Word workflow", instructions);
        Assert.DoesNotContain("PowerPoint workflow", instructions);
    }

    [Fact]
    public void WordOnlyCatalogExplainsPositionRefreshWithoutOtherHosts()
    {
        var instructions = OfficeMcpInstructions.Build(Tools("word_get_context", "word_read_document", "word_replace_range"));

        Assert.Contains("Word workflow", instructions);
        Assert.Contains("Character positions are snapshots", instructions);
        Assert.Contains("word_replace_range", instructions);
        Assert.DoesNotContain("Excel workflow", instructions);
        Assert.DoesNotContain("PowerPoint workflow", instructions);
    }

    [Fact]
    public void PowerPointOnlyCatalogExplainsShapeNamesAndSlideRefresh()
    {
        var instructions = OfficeMcpInstructions.Build(Tools("powerpoint_list_slides", "powerpoint_read_slide", "powerpoint_move_slide", "powerpoint_format_shape"));

        Assert.Contains("PowerPoint workflow", instructions);
        Assert.Contains("shape_name", instructions);
        Assert.Contains("Slide numbers are snapshots", instructions);
        Assert.Contains("powerpoint_format_shape", instructions);
        Assert.DoesNotContain("Excel workflow", instructions);
        Assert.DoesNotContain("Word workflow", instructions);
    }

    [Fact]
    public void MixedCatalogKeepsPreferredHostOrderAndDoesNotDuplicateSections()
    {
        var definitions = Tools(
            "powerpoint_get_context",
            "powerpoint_read_slide",
            "excel_get_context",
            "excel_read_range",
            "powerpoint_format_shape");

        var instructions = OfficeMcpInstructions.Build(definitions);

        Assert.True(instructions.IndexOf("PowerPoint workflow", StringComparison.Ordinal) < instructions.IndexOf("Excel workflow", StringComparison.Ordinal));
        Assert.Equal(1, Count(instructions, "PowerPoint workflow"));
        Assert.Equal(1, Count(instructions, "Excel workflow"));
        Assert.DoesNotContain("Word workflow", instructions);
    }

    [Fact]
    public void UnknownToolsReceiveGenericGuidanceWithoutInventedHostCalls()
    {
        var instructions = OfficeMcpInstructions.Build(Tools("onenote_read_page"));

        Assert.Contains("Additional connected Office tools", instructions);
        Assert.DoesNotContain("Excel workflow", instructions);
        Assert.DoesNotContain("Word workflow", instructions);
        Assert.DoesNotContain("PowerPoint workflow", instructions);
        Assert.DoesNotContain("Outlook workflow", instructions);
    }

    [Fact]
    public void OutlookOnlyCatalogExplainsMailboxDisciplineAndIrreversibleSend()
    {
        var definitions = Tools(
            "outlook_get_context",
            "outlook_list_folders",
            "outlook_list_items",
            "outlook_read_item",
            "outlook_create_draft",
            "outlook_update_draft",
            "outlook_delete_draft",
            "outlook_send_draft");
        definitions.Single(definition => definition.Name == "outlook_send_draft").Irreversible = true;

        var instructions = OfficeMcpInstructions.Build(definitions);

        Assert.Contains("Outlook workflow", instructions);
        Assert.Contains("no document checkpoints", instructions);
        Assert.Contains("outlook_list_folders", instructions);
        Assert.Contains("entry_id", instructions);
        Assert.Contains("irreversible", instructions);
        Assert.Contains("outlook_send_draft", instructions);
        Assert.DoesNotContain("Excel workflow", instructions);
        Assert.DoesNotContain("Word workflow", instructions);
        Assert.DoesNotContain("PowerPoint workflow", instructions);
    }

    [Fact]
    public void OutlookReadOnlyCatalogOmitsSendAndDraftGuidance()
    {
        var instructions = OfficeMcpInstructions.Build(Tools("outlook_get_context", "outlook_list_folders", "outlook_list_items", "outlook_read_item"));

        Assert.Contains("Outlook workflow", instructions);
        Assert.Contains("outlook_read_item", instructions);
        Assert.DoesNotContain("outlook_send_draft", instructions);
        Assert.DoesNotContain("outlook_create_draft", instructions);
        Assert.DoesNotContain("Compose email in stages", instructions);
    }

    [Fact]
    public void MixedExcelAndOutlookCatalogKeepsFirstSeenHostOrder()
    {
        var definitions = Tools(
            "outlook_get_context",
            "outlook_read_item",
            "excel_get_context",
            "excel_read_range");

        var instructions = OfficeMcpInstructions.Build(definitions);

        Assert.True(instructions.IndexOf("Outlook workflow", StringComparison.Ordinal) < instructions.IndexOf("Excel workflow", StringComparison.Ordinal));
        Assert.Equal(1, Count(instructions, "Outlook workflow"));
        Assert.Equal(1, Count(instructions, "Excel workflow"));
        Assert.DoesNotContain("Word workflow", instructions);
    }

    [Fact]
    public void IrreversibleToolsAreAlwaysCalledOut()
    {
        var withIrreversible = OfficeMcpInstructions.Build(
        [
            new OfficeToolDefinition { Name = "outlook_create_draft", Destructive = true },
            new OfficeToolDefinition { Name = "outlook_send_draft", Destructive = true, Irreversible = true }
        ]);
        var withoutIrreversible = OfficeMcpInstructions.Build(
        [
            new OfficeToolDefinition { Name = "outlook_create_draft", Destructive = true },
            new OfficeToolDefinition { Name = "outlook_update_draft", Destructive = true }
        ]);

        Assert.Contains("Irreversible actions that no Ribbon checkpoint can undo: outlook_send_draft.", withIrreversible);
        Assert.DoesNotContain("Irreversible actions", withoutIrreversible);
    }

    [Fact]
    public void ReadOnlyMetadataControlsWhetherMutationRefreshGuidanceAppears()
    {
        var readOnly = OfficeMcpInstructions.Build([new OfficeToolDefinition { Name = "word_custom_inspect", Destructive = false }]);
        var destructive = OfficeMcpInstructions.Build([new OfficeToolDefinition { Name = "word_custom_change", Destructive = true }]);

        Assert.DoesNotContain("Character positions are snapshots", readOnly);
        Assert.Contains("Character positions are snapshots", destructive);
    }

    private static List<OfficeToolDefinition> Tools(params string[] names)
    {
        return names.Select(name => new OfficeToolDefinition
        {
            Name = name,
            Destructive = !name.Contains("_get_", StringComparison.OrdinalIgnoreCase)
                && !name.Contains("_list_", StringComparison.OrdinalIgnoreCase)
                && !name.Contains("_read_", StringComparison.OrdinalIgnoreCase)
        }).ToList();
    }

    private static int Count(string value, string fragment)
    {
        var count = 0;
        var start = 0;
        while ((start = value.IndexOf(fragment, start, StringComparison.Ordinal)) >= 0)
        {
            count++;
            start += fragment.Length;
        }
        return count;
    }
}
