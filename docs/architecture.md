# Ribbon architecture

## Why three add-ins and one broker

VSTO add-ins run inside their Office host and can use the full COM object model. That is the control surface Ribbon needs, but an in-process add-in is a poor place to download runtimes, launch arbitrary agent processes, or own long-running protocol sessions.

Ribbon therefore separates host access from agent infrastructure:

- `Grid`, `Quill`, and `Deck` are independent VSTO deployment units because Office loads add-ins per application.
- `Ribbon.Vsto` makes those three add-ins behave like one product.
- `Ribbon.Broker` is a single per-user process shared by every running Office application.
- Every ACP agent receives one MCP server that can expose tools from all currently connected Office hosts.

This keeps failure and deployment boundaries aligned with Office while preserving a unified user experience.

## Connection sequence

1. An Office add-in starts and connects to the `Ribbon.Broker.v1` named pipe.
2. It registers a unique host id and its application kind.
3. The user installs and selects an agent in the shared task pane.
4. The broker launches the agent and performs ACP `initialize` over stdio.
5. The broker creates an isolated session directory and calls ACP `session/new`.
6. `session/new` includes `Ribbon.Broker.exe --mcp-stdio --host-id <id>` as a stdio MCP server.
7. The MCP proxy asks the primary broker for the live tool catalog. The preferred Office host is listed first, followed by other connected hosts.
8. MCP calls are routed to the VSTO process that owns the selected tool and executed on that Office application's UI thread.
9. ACP session updates stream back through the broker to the task pane.

The broker pipe uses a small versioned envelope from `Ribbon.Contracts`. Payloads are JSON strings so both .NET Framework 4.8 and modern .NET can share the protocol without sharing a JSON runtime.

## Process and thread boundaries

| Component | Runtime | Responsibility |
| --- | --- | --- |
| Office add-ins | .NET Framework 4.8 in Office | Task pane and COM operations |
| Shared VSTO library | .NET Framework 4.8 in Office | Broker lifecycle, pipe RPC, permissions, UI |
| Ribbon Broker | .NET 10 out of process | Registry, installation, ACP sessions, MCP routing |
| ACP agent | Agent-defined process | Reasoning and tool selection |
| MCP proxy | Additional Ribbon Broker process | stdio MCP endpoint supplied to one ACP session |

COM calls never run on the broker thread. Each host adapter dispatches work back to the synchronization context captured during VSTO startup.

## Registry and installation

Ribbon consumes the public ACP Registry document and supports these Windows distributions:

- `binary`: HTTPS ZIP download into `%LOCALAPPDATA%\Ribbon\agents`
- `npx`: existing Node.js or a Ribbon-managed current LTS runtime
- `uvx`: an existing `uvx` executable; managed uv provisioning is not implemented yet

Archive extraction rejects paths that escape the installation directory. Managed Node downloads are checked against Node.js `SHASUMS256.txt`. The ACP Registry does not currently provide artifact hashes for every binary agent, so Ribbon cannot independently checksum those downloads.

Installed-agent records are local and remain available when the Registry is offline. Registry refresh failures fall back to the cached document when one exists.

## Security model

- Broker named pipes are created with `CurrentUserOnly`.
- Only HTTPS registry and runtime downloads are accepted.
- ZIP entries are contained within a dedicated installation directory.
- Agent sessions run in Ribbon-owned directories rather than in the user's document folder.
- ACP filesystem and terminal client capabilities are declared unavailable.
- User-visible ACP permission requests are relayed to an Office modal confirmation dialog, defaulting to denial.
- MCP write tools are marked destructive; read tools are marked read-only.

An installed ACP agent is still executable code with the permissions of the signed-in user. Production releases should add publisher/signature information to the registry UI and sign Ribbon's own binaries and deployment manifests.

## Extending Office tools

Add a tool in the application-specific `IOfficeHost` implementation:

1. Return an `OfficeToolDefinition` with a globally unique name such as `excel_format_range`.
2. Provide a strict JSON Schema and set `Destructive` accurately.
3. Handle the name in `InvokeAsync`.
4. Execute COM access through the Office dispatcher.
5. Return JSON-safe structured data in `OfficeToolResult.ContentJson`.

No ACP or MCP changes are required. The MCP proxy builds its tool list dynamically from connected hosts.
