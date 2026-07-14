# Ribbon contributor guide

## Purpose

Ribbon embeds installable ACP agents in Microsoft Office. Excel, Word, and PowerPoint have separate thin VSTO add-ins, while one shared out-of-process broker owns the ACP Registry, agent processes, conversations, permissions, and the unified Office MCP server.

Keep that separation intact. Application-specific COM automation belongs in the relevant host adapter; agent and protocol infrastructure belongs in `Ribbon.Broker`; shared task-pane behavior belongs in `Ribbon.Vsto`.

## Project map

- `Ribbon.Contracts` — versioned JSON pipe contracts; targets `netstandard2.0` so both runtimes can consume it.
- `Ribbon.Broker` — .NET 8 broker, ACP client, Registry installer, session manager, and stdio MCP proxy.
- `Ribbon.Vsto` — .NET Framework 4.8 pipe client, shared task pane, agent manager, and permission UI.
- `Grid` — Excel VSTO host and Excel tools.
- `Quill` — Word VSTO host and Word tools.
- `Deck` — PowerPoint VSTO host and PowerPoint tools.
- `docs/architecture.md` — protocol, process, security, and extension details.

The legacy Grid chat/MCP source directories remain for reference but are intentionally excluded from `Grid.csproj`. Do not reconnect them to the build unless the architecture is deliberately being reconsidered.

## Architecture invariants

1. Office COM calls stay inside the owning Office process and execute on the captured Office UI synchronization context.
2. VSTO add-ins do not download runtimes or directly launch ACP agents. They communicate with the broker through the current-user named pipe.
3. ACP agents receive `Ribbon.Broker.exe --mcp-stdio` as their Office MCP server. Do not expose the VSTO pipe directly to arbitrary agent processes.
4. Tool names are globally unique and application-prefixed: `excel_*`, `word_*`, or `powerpoint_*`.
5. Tool schemas must be strict JSON Schema objects. Mark every user-visible mutation as destructive.
6. Keep the broker wire protocol backward-conscious. Increment `RibbonProtocol.Version` for an incompatible envelope or payload change.
7. Preserve the per-user process and storage model under `%LOCALAPPDATA%\Ribbon`; never require administrator privileges for normal use.

## Coding guidance

- `Ribbon.Vsto` and the Office projects must remain compatible with .NET Framework 4.8 and C# 9.
- `Ribbon.Broker` targets .NET 8 for current process, JSON, pipe, and archive APIs.
- Avoid adding a JSON package to the VSTO projects. `Ribbon.Vsto.JsonCodec` deliberately uses the framework serializer, while the broker uses `System.Text.Json`.
- Return structured, JSON-safe values from Office tools. Convert COM-specific values before crossing the pipe.
- Catch exceptions at process and protocol boundaries and return useful errors. Do not silently swallow tool failures.
- Release temporary COM objects where practical, but do not aggressively final-release objects owned by Office such as the application or active selection.
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
dotnet build Ribbon.Vsto\Ribbon.Vsto.csproj --no-restore
```

For changes to routing, verify at least MCP `initialize`, `tools/list`, and `tools/call` through a connected synthetic or real Office host. Do not install a third-party ACP agent merely to run a routine test unless that external state change is explicitly intended.

When changing deployment content, inspect each generated VSTO application manifest and confirm that `Ribbon.Broker.exe`, `.dll`, `.deps.json`, and `.runtimeconfig.json` remain included.

## Change hygiene

- Do not edit generated VSTO designer files unless the corresponding designer contract requires it.
- Do not commit `bin`, `obj`, restored `packages`, `.vs`, user files, or temporary signing keys.
- Preserve unrelated local changes.
- Update `README.md` and `docs/architecture.md` when behavior, prerequisites, security boundaries, or supported distributions change.
