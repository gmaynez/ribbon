using System.Text;
using Ribbon.Contracts;

namespace Ribbon.Broker.Mcp;

internal static class OfficeMcpInstructions
{
    public static string Build(IEnumerable<OfficeToolDefinition>? definitions)
    {
        var orderedDefinitions = (definitions ?? [])
            .Where(definition => !string.IsNullOrWhiteSpace(definition.Name))
            .GroupBy(definition => definition.Name.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();
        var orderedNames = orderedDefinitions.Select(definition => definition.Name.Trim()).ToList();
        var names = orderedNames.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var destructiveNames = orderedDefinitions
            .Where(definition => definition.Destructive)
            .Select(definition => definition.Name.Trim())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var irreversibleNames = orderedDefinitions
            .Where(definition => definition.Irreversible)
            .Select(definition => definition.Name.Trim())
            .ToList();
        var instructions = new StringBuilder();
        instructions.AppendLine("Use Ribbon to work directly in the user's currently connected Microsoft Office documents. Choose only tools present in tools/list; never invent an unavailable tool or parameter. Treat every tool description and input schema as authoritative, including for capabilities added after initialization.");
        instructions.AppendLine("Work naturally in a short inspect → act → verify loop: inspect only enough context to identify the target, make the smallest task-oriented change that satisfies the request, then read back the affected state. Treat identifiers and resolved locations returned by tools as authoritative. If the target or destructive intent remains ambiguous after inspection, ask the user instead of guessing. After an error, use its details to adjust the next call rather than repeating the same call unchanged. In the final response, summarize visible changes and disclose anything that could not be verified.");

        if (orderedNames.Count == 0)
        {
            instructions.Append("No Office tools are currently available. Do not fabricate tool calls; tell the user to open or reconnect a supported Office document and retry.");
            return instructions.ToString();
        }

        foreach (var host in HostOrder(orderedNames))
        {
            switch (host)
            {
                case "excel":
                    AppendExcel(instructions, names, destructiveNames);
                    break;
                case "word":
                    AppendWord(instructions, names, destructiveNames);
                    break;
                case "powerpoint":
                    AppendPowerPoint(instructions, names, destructiveNames);
                    break;
                case "outlook":
                    AppendOutlook(instructions, names, destructiveNames);
                    break;
            }
        }

        if (irreversibleNames.Count > 0)
        {
            instructions.AppendLine();
            instructions.Append("Irreversible actions that no Ribbon checkpoint can undo: "
                + string.Join(", ", irreversibleNames) + ". Call them only when the request explicitly asks for the outcome, after verifying the affected state.");

        }
        else if (orderedNames.Any(name => HostPrefix(name) == null))
        {
            instructions.AppendLine();
            instructions.Append("Additional connected Office tools are available. Follow each tool's description and strict input schema, and apply the same inspect → act → verify discipline.");
        }

        return instructions.ToString().TrimEnd();
    }

    private static IEnumerable<string> HostOrder(IEnumerable<string> names)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in names)
        {
            var host = HostPrefix(name);
            if (host != null && seen.Add(host)) yield return host;
        }
    }

    private static string? HostPrefix(string name)
    {
        if (name.StartsWith("excel_", StringComparison.OrdinalIgnoreCase)) return "excel";
        if (name.StartsWith("word_", StringComparison.OrdinalIgnoreCase)) return "word";
        if (name.StartsWith("powerpoint_", StringComparison.OrdinalIgnoreCase)) return "powerpoint";
        if (name.StartsWith("outlook_", StringComparison.OrdinalIgnoreCase)) return "outlook";
        return null;
    }

    private static void AppendExcel(StringBuilder instructions, ISet<string> names, ISet<string> destructiveNames)
    {
        instructions.AppendLine();
        instructions.AppendLine("Excel workflow (available excel_* tools):");
        instructions.AppendLine("- Use the descriptions and schemas for any additional available Excel tools not called out below.");
        if (Has(names, "excel_get_context")) instructions.AppendLine("- Call excel_get_context first when the workbook, worksheet, selection, or active cell is not explicit.");
        if (Has(names, "excel_list_sheets")) instructions.AppendLine("- Use excel_list_sheets to discover worksheet names and used ranges instead of guessing sheet names.");
        if (Has(names, "excel_read_range")) instructions.AppendLine("- Use excel_read_range for a targeted source or verification read; keep the range as small as the task permits.");
        if (Has(names, "excel_write_range")) instructions.AppendLine("- Use excel_write_range only for literal values. Formula-looking strings remain literal by design.");
        if (Has(names, "excel_write_formulas")) instructions.AppendLine("- Use excel_write_formulas for formulas, and supply explicit formulas beginning with '='.");
        if (Has(names, "excel_clear_range")) instructions.AppendLine("- Use excel_clear_range with the narrowest clear mode that matches the request.");
        if (Has(names, "excel_format_range")) instructions.AppendLine("- Treat excel_format_range as a patch: specify only properties the user wants changed so existing styling is preserved.");
        if (Has(names, "excel_create_table")) instructions.AppendLine("- Before excel_create_table, inspect the complete source range and confirm whether it already contains headers.");
        if (Has(names, "excel_create_chart")) instructions.AppendLine("- Before excel_create_chart, inspect the source range and choose a chart type and placement that match the data.");
        if (Has(names, "excel_add_sheet")) instructions.AppendLine("- Use excel_add_sheet only after checking existing sheet names and choosing an unambiguous position.");
        if (Has(names, "excel_read_range") && destructiveNames.Any(name => name.StartsWith("excel_", StringComparison.OrdinalIgnoreCase)))
            instructions.AppendLine("- After a mutation, use the returned workbook, worksheet, address, and dimensions for a focused excel_read_range verification.");
    }

    private static void AppendWord(StringBuilder instructions, ISet<string> names, ISet<string> destructiveNames)
    {
        instructions.AppendLine();
        instructions.AppendLine("Word workflow (available word_* tools):");
        instructions.AppendLine("- Use the descriptions and schemas for any additional available Word tools not called out below.");
        if (Has(names, "word_get_context")) instructions.AppendLine("- Call word_get_context first when the document, selection, or current story is not explicit.");
        if (Has(names, "word_list_headings")) instructions.AppendLine("- Use word_list_headings to understand long-document structure before choosing where to read or insert.");
        if (Has(names, "word_read_document")) instructions.AppendLine("- Use word_read_document for a bounded text slice and take exact character positions from its result.");
        if (Has(names, "word_replace_selection")) instructions.AppendLine("- Use word_replace_selection only when the user's current selection is intentionally the target.");
        if (Has(names, "word_replace_range")) instructions.AppendLine("- Use word_replace_range for a precise span obtained from a recent read; an empty replacement deletes that span.");
        if (Has(names, "word_find_replace")) instructions.AppendLine("- Use word_find_replace for bounded literal replacement, with case and whole-word behavior chosen deliberately.");
        if (Has(names, "word_insert_text") || Has(names, "word_append_text")) instructions.AppendLine("- Choose the insertion tool and position that express the requested document location; do not simulate structure with arbitrary spacing.");
        if (Has(names, "word_insert_heading") || Has(names, "word_insert_list") || Has(names, "word_insert_table") || Has(names, "word_insert_page_break"))
            instructions.AppendLine("- Prefer the dedicated heading, list, table, and page-break tools for document structure when they are available.");
        if (Has(names, "word_format_range")) instructions.AppendLine("- Treat word_format_range as a patch and specify only the style, font, paragraph, or highlight properties that should change.");
        if (Has(names, "word_add_comment")) instructions.AppendLine("- Use word_add_comment for review feedback; do not insert comment text into the document body.");
        if (destructiveNames.Any(name => name.StartsWith("word_", StringComparison.OrdinalIgnoreCase)))
            instructions.AppendLine("- Character positions are snapshots and move after edits. Refresh context, headings, or document text before a later position-based mutation, then verify the affected text or structure.");
    }

    private static void AppendPowerPoint(StringBuilder instructions, ISet<string> names, ISet<string> destructiveNames)
    {
        instructions.AppendLine();
        instructions.AppendLine("PowerPoint workflow (available powerpoint_* tools):");
        instructions.AppendLine("- Use the descriptions and schemas for any additional available PowerPoint tools not called out below.");
        if (Has(names, "powerpoint_get_context")) instructions.AppendLine("- Call powerpoint_get_context first when the presentation, selected slide, page size, or selected shapes are not explicit.");
        if (Has(names, "powerpoint_list_slides")) instructions.AppendLine("- Use powerpoint_list_slides to understand the deck outline and obtain current one-based slide numbers.");
        if (Has(names, "powerpoint_read_slide")) instructions.AppendLine("- Use powerpoint_read_slide before editing a slide and use returned shape_name values for later shape mutations.");
        if (Has(names, "powerpoint_add_slide") || Has(names, "powerpoint_duplicate_slide") || Has(names, "powerpoint_move_slide") || Has(names, "powerpoint_delete_slide"))
            instructions.AppendLine("- Slide numbers are snapshots. After adding, duplicating, moving, or deleting slides, refresh the outline before another slide-number-based call.");
        if (Has(names, "powerpoint_add_textbox") || Has(names, "powerpoint_add_shape")) instructions.AppendLine("- Place text boxes and diagram shapes using point-based geometry derived from the presentation page size and nearby slide content.");
        if (Has(names, "powerpoint_format_shape")) instructions.AppendLine("- Treat powerpoint_format_shape as a patch: change only the requested geometry, appearance, text, or z-order properties.");
        if (Has(names, "powerpoint_delete_shape")) instructions.AppendLine("- Delete a shape only by a shape_name obtained from a recent slide read or creation result.");
        if (Has(names, "powerpoint_add_table")) instructions.AppendLine("- Use powerpoint_add_table for structured row-and-column content rather than aligning text manually.");
        if (Has(names, "powerpoint_add_chart")) instructions.AppendLine("- Use powerpoint_add_chart for numeric categories and series; choose the chart type, title, legend, and placement to fit the slide's message.");
        if (Has(names, "powerpoint_add_image")) instructions.AppendLine("- powerpoint_add_image accepts an existing absolute local path; do not pass a URL or assume Ribbon downloads the image.");
        if (Has(names, "powerpoint_set_speaker_notes")) instructions.AppendLine("- Put presenter-only guidance in powerpoint_set_speaker_notes, not in visible slide text.");
        if (Has(names, "powerpoint_set_slide_background")) instructions.AppendLine("- Use powerpoint_set_slide_background only when changing that slide independently from its master is intentional.");
        if (Has(names, "powerpoint_find_replace")) instructions.AppendLine("- Use powerpoint_find_replace for bounded literal text changes and opt into speaker notes only when the request includes them.");
        if (Has(names, "powerpoint_read_slide") && destructiveNames.Any(name => name.StartsWith("powerpoint_", StringComparison.OrdinalIgnoreCase)))
            instructions.AppendLine("- After a mutation, read the affected slide and verify its title, shape text, geometry, table/chart presence, notes, or other changed state.");
    }

    private static void AppendOutlook(StringBuilder instructions, ISet<string> names, ISet<string> destructiveNames)
    {
        instructions.AppendLine();
        instructions.AppendLine("Outlook workflow (available outlook_* tools):");
        instructions.AppendLine("- Use the descriptions and schemas for any additional available Outlook tools not called out below.");
        instructions.AppendLine("- Outlook has no document checkpoints. Treat reads as the only source of truth and prefer the least invasive action that satisfies the request.");
        if (Has(names, "outlook_get_context")) instructions.AppendLine("- Call outlook_get_context first to identify the mailbox, active folder, and selection before touching any mail.");
        if (Has(names, "outlook_list_folders")) instructions.AppendLine("- Use outlook_list_folders to discover real folder paths instead of guessing folder names.");
        if (Has(names, "outlook_list_items")) instructions.AppendLine("- Use outlook_list_items for a bounded folder listing and take entry_id values from its result for later reads.");
        if (Has(names, "outlook_read_item")) instructions.AppendLine("- Use outlook_read_item for one item's bounded content; do not bulk-read entire folders when a listing plus a few reads suffice.");
        if (Has(names, "outlook_create_draft") || Has(names, "outlook_update_draft"))
            instructions.AppendLine("- Compose email in stages: create a draft, patch it, then read it back to verify recipients, subject, and body before anything is sent.");
        if (Has(names, "outlook_send_draft"))
            instructions.AppendLine("- outlook_send_draft is irreversible: it delivers real mail and no checkpoint can undo it. Send only when the user explicitly asked to send, and only after verifying the draft.");
        if (Has(names, "outlook_delete_draft")) instructions.AppendLine("- Use outlook_delete_draft only for drafts the agent created or that the user explicitly asked to remove.");
    }

    private static bool Has(ISet<string> names, string name) => names.Contains(name);
}
