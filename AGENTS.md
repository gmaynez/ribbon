# Ribbon contributor guide

## Purpose

Ribbon embeds installable ACP agents in Microsoft Office. Excel, Word, and PowerPoint have separate thin VSTO add-ins, while one shared out-of-process broker owns the ACP Registry, agent processes, conversations, permissions, and the unified Office MCP server.

Keep that separation intact. Application-specific COM automation belongs in the relevant host adapter; agent and protocol infrastructure belongs in `Ribbon.Broker`; shared task-pane behavior belongs in `Ribbon.Vsto`.

## Project map

- `Ribbon.Contracts` — versioned JSON pipe contracts; targets `netstandard2.0` so both runtimes can consume it.
- `Ribbon.Broker` — .NET 10 broker, ACP client, Registry installer, session manager, and stdio MCP proxy.
- `Ribbon.Broker.Tests` — .NET 10 broker tests, including capability-aware MCP instruction composition.
- `Ribbon.Vsto` — .NET Framework 4.8 pipe client, shared task pane, conversation history, agent manager, and permission UI. `ConversationHistory.cs` owns persistent records; `ConversationHistoryDialog.cs` owns browsing and deletion.
- `Grid` — Excel VSTO host and Excel tools.
- `Quill` — Word VSTO host and Word tools.
- `Deck` — PowerPoint VSTO host and PowerPoint tools.
- `docs/architecture.md` — protocol, process, security, and extension details.

The active Excel tool surface is split deliberately:

- `Grid/Office/GridOfficeHost.cs` — MCP-visible definitions and invocation routing.
- `Grid/Office/ExcelToolSchemas.cs` — strict input schemas and agent-facing parameter descriptions.
- `Grid/Office/ExcelToolModels.cs` — JSON request DTOs compatible with `JavaScriptSerializer`.
- `Grid/Office/ExcelAutomationService.cs` — Office-thread validation, COM operations, and JSON-safe results.
- `Grid/Office/ExcelCheckpointService.cs` — turn snapshots and in-place restoration for the Excel tool surface.

`Grid.csproj` uses explicit `Compile` items rather than SDK-style globbing. Add every new active source file to the project.

The active Word tool surface follows the same split:

- `Quill/Office/QuillOfficeHost.cs` — MCP-visible definitions and invocation routing.
- `Quill/Office/WordToolSchemas.cs` — strict schemas and agent-facing parameter descriptions.
- `Quill/Office/WordToolModels.cs` — `JavaScriptSerializer` request DTOs.
- `Quill/Office/WordAutomationService.cs` — Office-thread validation and Word COM operations.
- `Quill/Office/WordCheckpointService.cs` — main-story WordOpenXML checkpoints and restoration.

`Quill.csproj` includes `Office/*.cs`, so new Word tool source files are compiled automatically.

The active PowerPoint tool surface follows the same split:

- `Deck/Office/DeckOfficeHost.cs` — MCP-visible definitions and invocation routing.
- `Deck/Office/PowerPointToolSchemas.cs` — strict input schemas and agent-facing parameter descriptions.
- `Deck/Office/PowerPointToolModels.cs` — `JavaScriptSerializer` request DTOs.
- `Deck/Office/PowerPointAutomationService.cs` — Office-thread validation and PowerPoint COM operations.
- `Deck/Office/PowerPointCheckpointService.cs` — presentation-copy checkpoints and slide restoration.

`Deck.csproj` includes `Office/*.cs`, so new PowerPoint tool source files are compiled automatically.

## Architecture invariants

1. Office COM calls stay inside the owning Office process and execute on the captured Office UI synchronization context.
2. VSTO add-ins do not download runtimes or directly launch ACP agents. They communicate with the broker through the current-user named pipe.
3. ACP agents receive `Ribbon.Broker.exe --mcp-stdio` as their Office MCP server. Do not expose the VSTO pipe directly to arbitrary agent processes.
4. Tool names are globally unique and application-prefixed: `excel_*`, `word_*`, or `powerpoint_*`.
5. Tool schemas must be strict JSON Schema objects. Mark every user-visible mutation as destructive.
6. Keep the broker wire protocol backward-conscious. Increment `RibbonProtocol.Version` for an incompatible envelope or payload change.
7. Preserve the per-user process and storage model under `%LOCALAPPDATA%\Ribbon`; never require administrator privileges for normal use.
8. Keep Office tools task-oriented rather than exposing arbitrary COM dispatch. Tool descriptions and results should help an agent decide its next safe action.
9. Excel range operations accept one contiguous A1 range. Keep reads and writes bounded to 100,000 cells per call so Office's UI thread and the agent context remain responsive.
10. Keep literal Excel values and formulas as separate operations. `excel_write_range` must not execute formula-like text; formulas belong in `excel_write_formulas` and must begin with `=`.
11. Treat `excel_format_range` as a patch: omitted formatting properties preserve the workbook's current styling.
12. Word main-story ranges use zero-based character positions and reads are bounded to 200,000 characters. Positions shift after text mutations, so agents must refresh before later position-based edits.
13. Treat `word_format_range` as a patch. Keep structured insertions such as headings, lists, tables, comments, and page breaks as task-oriented tools rather than raw COM access.
14. PowerPoint slide numbers are one-based snapshots of the current presentation order. Refresh them after slide lifecycle changes and use the returned `shape_name` for later shape mutations.
15. Treat `powerpoint_format_shape` as a patch. Express geometry in points, keep chart and table payloads bounded to 10,000 data cells, and accept only existing absolute local paths for image insertion.
16. Compose MCP server instructions from the live tool catalog. Mention only available tool names, preserve preferred-host order, keep a safe no-tools fallback, and test single-host, mixed-host, partial, and unknown-host catalogs.
17. Create and restore turn checkpoints inside the owning Office process. Restore only the documented Ribbon tool surface, capture a safety checkpoint first, and reset the ACP session after restore so the agent must inspect fresh state.
18. Keep Ribbon's local conversation record authoritative for display. Gate ACP `session/list`, `session/resume`, and `session/load` by advertised capabilities, reuse only Ribbon-owned session directories, bind continuation to the original Office document, and make cross-document history read-only.

## Coding guidance

- `Ribbon.Vsto` and the Office projects must remain compatible with .NET Framework 4.8 and C# 9.
- `Ribbon.Broker` targets .NET 10 with C# 14 for current process, JSON, pipe, and archive APIs.
- Avoid adding a JSON package to the VSTO projects. `Ribbon.Vsto.JsonCodec` deliberately uses the framework serializer, while the broker uses `System.Text.Json`.
- Return structured, JSON-safe values from Office tools. Convert COM-specific values before crossing the pipe.
- Resolve and return the actual workbook, worksheet, and range after mutations; include affected dimensions for matrix-shaped writes so agents can verify with a targeted follow-up read.
- Catch exceptions at process and protocol boundaries and return useful errors. Do not silently swallow tool failures.
- Release temporary COM objects where practical. Use a balanced `Marshal.ReleaseComObject` for temporary RCWs; never `FinalReleaseComObject` on application-owned objects such as the application, active workbook, active worksheet, or selection.
- Keep UI work small and non-blocking. Network, installation, and agent work belongs in the broker.

## Security expectations

- Named-pipe servers must remain current-user-only.
- Registry and runtime downloads must use HTTPS.
- Continue validating archive extraction paths against traversal.
- Verify checksums whenever an authoritative checksum is available.
- Default permission prompts to denial and never auto-select an allow choice without user confirmation.
- Do not commit certificates, private keys, credentials, downloaded agents, runtime archives, or `%LOCALAPPDATA%\Ribbon` state.

## Build and verification

Run the full build with Visual Studio MSBuild because the solution contains VSTO projects:

```powershell
& 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe' Ribbon.slnx /t:Build /p:Configuration=Debug /m /verbosity:minimal
```

The broker and shared libraries can also be checked independently:

```powershell
dotnet build Ribbon.Broker\Ribbon.Broker.csproj --no-restore
dotnet test Ribbon.Broker.Tests\Ribbon.Broker.Tests.csproj --no-restore
dotnet build Ribbon.Vsto\Ribbon.Vsto.csproj --no-restore
```

For changes to routing, verify at least MCP `initialize`, `tools/list`, and `tools/call` through a connected synthetic or real Office host. Do not install a third-party ACP agent merely to run a routine test unless that external state change is explicitly intended.

For conversation-history changes, cover capability combinations for ACP load, resume, and list; metadata title updates; atomic local serialization; current-versus-other-document filtering; multiple unsaved documents in one Office process; fresh-continuation disclosure; narrow task-pane layout; and history-dialog layout. Never use or delete the user's real `%LOCALAPPDATA%\Ribbon\Conversations` data in an automated test.

For Excel tool changes:

- Confirm every schema parses as JSON and keeps `additionalProperties: false`, including nested objects.
- Exercise definitions and dispatch through `GridOfficeHost.GetTools` and `GridOfficeHost.InvokeAsync`, not only by calling the automation service directly.
- Use a fresh disposable workbook in a separate hidden Excel instance for integration checks; close it without saving and never repurpose the user's open workbook as test data.
- Cover representative values, formulas, formatting, tables or charts when touched, and read back the resulting workbook state.
- Re-run the complete Release build after Excel integration tests.

For Word tool changes:

- Validate every schema as JSON with strict nested objects.
- Exercise definitions and dispatch through `QuillOfficeHost.GetTools` and `QuillOfficeHost.InvokeAsync`.
- Use a fresh unsaved document in a separate hidden Word instance; close it without saving and never use the user's open document as test data.
- Read back headings and text, and inspect the real Word object model for formatting, lists, tables, comments, or breaks touched by the change.
- Re-run the complete Release build and inspect Quill's deployment manifest.

For PowerPoint tool changes:

- Confirm every schema parses as JSON and keeps `additionalProperties: false`, including nested objects.
- Exercise definitions and dispatch through `DeckOfficeHost.GetTools` and `DeckOfficeHost.InvokeAsync`, not only by calling the automation service directly.
- Use a fresh disposable presentation in a separate PowerPoint instance for integration checks; close it without saving and never repurpose the user's open presentation as test data.
- Cover slide lifecycle, structured shape reads, geometry and format patches, tables, native charts, images, speaker notes, backgrounds, and text replacement when touched.
- Re-run the complete Release build and inspect Deck's deployment manifest.

An open Office process can lock Debug output DLLs. Prefer a Release verification build while the user's Office session is open; ask before closing or terminating their application merely to unblock a build.

When changing deployment content, inspect each generated VSTO application manifest and confirm that `Ribbon.Broker.exe`, `.dll`, `.deps.json`, and `.runtimeconfig.json` remain included.

## Change hygiene

- Do not edit generated VSTO designer files unless the corresponding designer contract requires it.
- Do not commit `bin`, `obj`, restored `packages`, `.vs`, user files, or temporary signing keys.
- Preserve unrelated local changes.
- Update `README.md` and `docs/architecture.md` when behavior, prerequisites, security boundaries, or supported distributions change.
