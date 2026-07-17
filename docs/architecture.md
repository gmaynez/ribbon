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

1. An Office add-in starts and connects to the named pipe for its broker protocol version (`Ribbon.Broker.v1` for the current protocol).
2. It registers a unique host id and its application kind.
3. The user installs and selects an agent in the shared task pane.
4. The broker launches the agent and performs ACP `initialize` over stdio.
5. The broker creates an isolated session directory and calls ACP `session/new`.
6. `session/new` includes `Ribbon.Broker.exe --mcp-stdio --host-id <id>` as a stdio MCP server.
7. The task pane renders ACP `configOptions` with category `model` and changes them through `session/set_config_option`.
8. During MCP `initialize`, the proxy asks the primary broker for the live tool catalog. It composes an inspect–act–verify playbook from only those capabilities, with the preferred Office host first.
9. A later MCP `tools/list` refreshes the live catalog; every listed tool retains its host routing identity, description, strict schema, and mutation annotations.
10. Before each prompt, the owning VSTO host captures a local checkpoint of the document state that Ribbon tools can mutate.
11. MCP calls are routed to the VSTO process that owns the selected tool. Destructive definitions require user approval there before execution on that Office application's UI thread.
12. ACP message chunks, thoughts, plans, tool progress, and configuration updates stream back through the broker to the task pane.
13. Ribbon persists the structured transcript and document/session metadata locally. Reopening a saved conversation uses ACP `session/list` for advisory discovery and `session/resume` or `session/load` only when the agent advertises the corresponding capability.

The broker pipe uses a small versioned envelope from `Ribbon.Contracts`. Its pipe name is versioned as well, preventing a newly built add-in from silently attaching to an older long-lived broker that cannot provide newly required payloads. Payloads are JSON strings so both .NET Framework 4.8 and modern .NET can share the protocol without sharing a JSON runtime.

## Process and thread boundaries

| Component | Runtime | Responsibility |
| --- | --- | --- |
| Office add-ins | .NET Framework 4.8 in Office | Task pane and COM operations |
| Shared VSTO library | .NET Framework 4.8 in Office | Broker lifecycle, pipe RPC, permissions, UI |
| Ribbon Broker | .NET 10 out of process | Registry, installation, ACP sessions, MCP routing |
| ACP agent | Agent-defined process | Reasoning and tool selection |
| MCP proxy | Additional Ribbon Broker process | stdio MCP endpoint supplied to one ACP session |

One primary Ribbon Broker process is expected while at least one Office host is connected. Each live ACP session can add a second Ribbon Broker process running with `--mcp-stdio`; it is a thin stdio proxy, not another primary broker. When the user switches agents or restores a checkpoint, Ribbon closes the superseded session through capability-gated ACP `session/close`; agents without that capability have their now-unused runtime retired so its MCP proxy cannot accumulate. When an Office pipe client disconnects, the broker releases that client's sessions, cancels pending permission requests, and terminates agent runtimes that no longer serve a connected session. The primary broker exits after the last registered Office host disconnects, so the next Office launch starts the current deployed broker build.

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
- User-visible ACP permission requests preserve the agent's `allow_once`, `allow_always`, `reject_once`, and `reject_always` choices and default to denial.
- MCP write tools are marked destructive and are independently gated in the owning VSTO host; a remembered approval is scoped to one tool action and one active agent session.
- Checkpoint restore always requires explicit confirmation, captures the current state first, and resets the ACP session after the document changes.
- Conversation records contain transcript text and session metadata but never permission decisions, authentication credentials, or raw ACP payloads. Saved ACP working directories are accepted only when they remain under Ribbon's per-user session root.

## Conversation history

Ribbon treats its local history as the reliable UI record because ACP agents vary in persistence support. Each conversation is stored as structured text segments under `%LOCALAPPDATA%\Ribbon\Conversations`, so it can be rendered in the current Office light or dark theme rather than persisting theme-specific RTF colors. Records include the agent, selected model, ACP session id and original working directory, timestamps, and the Office document identity. The store uses atomic replacement, ignores damaged entries without blocking Office startup, and retains the newest 200 conversations.

The task pane provides **New** and **History** actions. History defaults to the active document and can optionally show every Ribbon conversation. Saved documents match by normalized path across Office restarts; inside a live Office process, Ribbon hashes the process and owning document-window handle so simultaneous unsaved documents remain distinct and the same document survives Save or Save As. Office's unique in-application name is the conservative fallback when a document has no window. If the active document changes during a chat, Ribbon requires a new conversation before another prompt can run.

Native continuity is capability-gated according to ACP v1:

- `sessionCapabilities.list` allows advisory discovery through paginated `session/list`.
- `sessionCapabilities.resume` reconnects the known session without replay. Ribbon supplies the original ACP working directory and a new MCP proxy targeting the current Office host.
- `loadSession` allows `session/load` when resume is unavailable. The broker consumes the required replay while the task pane renders its richer local transcript, then resumes live updates.
- `session_info_update` refreshes an agent-generated title in Ribbon's local record.

When neither restore method is supported, or the agent no longer knows the session, the transcript opens read-only and the user can explicitly create a fresh continuation. A fresh continuation copies the visible transcript for the user but states that the new agent session cannot see the old context. Conversations belonging to another document always open read-only; Ribbon never silently attaches their ACP context to the active document.

Conversation history and turn checkpoints remain deliberately separate. History persists chat and optional agent context across Office restarts. Checkpoints are temporary snapshots of the current Office process and are removed at host shutdown; reopening a chat does not resurrect old document checkpoints.

## Turn checkpoints

Checkpoints stay inside the Office trust boundary under `%LOCALAPPDATA%\Ribbon\Checkpoints\<host-id>` and are removed when that host runtime shuts down. The shared task pane keeps the newest twelve checkpoints and creates one before every prompt. Restore does not overwrite the user's saved file or replace the open Office document object; each host adapter restores the surface Ribbon can mutate into that existing object, then starts a fresh ACP session so later tool calls must inspect the restored state again.

- Excel checkpoints use `Workbook.SaveCopyAs`; restore copies the snapshot's sheets back into the active workbook as a group, preserving sheet content, formatting, tables, charts, and cross-sheet formulas.
- Word checkpoints store the main-story `WordOpenXML` and replace that story on restore, covering the text, formatting, lists, tables, comments, and breaks exposed by Ribbon's Word tools.
- PowerPoint checkpoints use `Presentation.SaveCopyAs`; restore replaces the active presentation's slides from the snapshot, including their shapes, tables, charts, notes, and slide formatting.

These are Ribbon-tool checkpoints rather than general Office version history. Changes outside the surfaces above, such as VBA projects or Word headers and footers, are not promised to roll back.

An installed ACP agent is still executable code with the permissions of the signed-in user. Production releases should add publisher/signature information to the registry UI and sign Ribbon's own binaries and deployment manifests.

## Extending Office tools

Add a tool in the application-specific `IOfficeHost` implementation:

1. Return an `OfficeToolDefinition` with a globally unique name such as `excel_format_range`.
2. Provide a strict JSON Schema and set `Destructive` accurately.
3. Handle the name in `InvokeAsync`.
4. Execute COM access through the Office dispatcher.
5. Return JSON-safe structured data in `OfficeToolResult.ContentJson`.

No ACP or MCP changes are required. The MCP proxy builds its tool list dynamically from connected hosts.

### Excel tool design

Excel tools are designed around agent tasks rather than exposing arbitrary COM dispatch:

- Reads return bounded, rectangular, JSON-safe matrices and identify the resolved workbook, sheet, and A1 address.
- Literal values and formulas use separate tools so intent is explicit and formula-like text is not executed accidentally.
- `excel_format_range` is patch-like: omitted properties preserve the workbook's current styling.
- Tables and charts consume an existing inspected range instead of accepting an opaque series of COM operations.
- Mutations return the resolved address and affected dimensions so an agent can verify its work with a targeted follow-up read.
- One read or write is limited to 100,000 cells to protect the Office UI thread and the agent context window.

### Word tool design

Word tools address the main document story by zero-based character positions, while selection-based operations can also act on the user's current Word story:

- Reads are bounded to 200,000 characters and return the resolved start and end positions.
- Heading discovery provides a structural index before agents edit long documents.
- Text insertion, headings, lists, tables, comments, and page breaks are separate task-oriented operations rather than arbitrary COM dispatch.
- `word_format_range` is patch-like: omitted style, font, paragraph, and highlight properties remain unchanged.
- Positions are snapshots and can shift after every text mutation; agents are instructed to refresh context, headings, or document text before a later position-based change.
- Tables are limited to 10,000 cells and 63 columns to protect Word's UI thread.

### PowerPoint tool design

`Deck` exposes presentation work as task-oriented `powerpoint_*` tools rather than arbitrary COM dispatch. The active surface supports bounded presentation outlines and structured slide reads; slide creation, duplication, movement, and deletion; title, text-box, diagram-shape, image, table, and native-chart authoring; patch-style shape formatting; speaker notes and solid backgrounds; and bounded literal find/replace.

Slide numbers are one-based and reflect the current presentation order, so agents should refresh context after slide lifecycle changes. Shape mutations use the `shape_name` returned by slide reads and creation results. Geometry is expressed in PowerPoint points. Structured reads return shape identity, type, geometry, text, table values, and chart/table flags so an agent can verify a mutation with a targeted follow-up read.

The native chart tool accepts a bounded category vector and equally sized numeric series. It assigns native PowerPoint chart series directly, avoiding an embedded-workbook activation that would launch Excel from the PowerPoint process. Image insertion accepts only an existing absolute local path and never performs network work. All authoring tools are destructive and all COM work runs through the captured PowerPoint UI synchronization context.
