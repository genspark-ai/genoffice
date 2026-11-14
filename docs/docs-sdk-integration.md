# Docs SDK integration and compatibility

ForaOffice separates the Docs editor from its desktop host and built-in agent.
The [SDK guide](../apps/docs/sdk/README.md) is the authoritative public API and
integration reference. This page describes responsibility and compatibility;
it does not declare the development SDK fully accepted.

## Components and responsibility

| Component              | Owns                                                                                                                       | Does not supply automatically                                         |
| ---------------------- | -------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------- |
| Docs editor            | Document editing, selection, review, live context, runtime tools, inspection and view annotations                          | Application storage, native windows, model inference or conversations |
| Host adapter           | File or managed-document storage, permissions, dialogs, clipboard, assets, fonts, printing/export and application settings | Permission to arbitrary files or a model configuration                |
| External Harness/agent | Conversation, model/provider selection, task orchestration and optional instruction resolution                             | Authority to a replacement document or a different selection          |
| Portable Node services | Authorized file processing, bounded Office CLI/resource helpers, shared exports and MCP adapters                           | An installed Electron/Qt runtime or a remote service deployment       |

The inherited Electron desktop agent is one consumer of the editor. Embedded
hosts do not need Genspark authentication, a provider key, the native AI sidebar
or a GenOffice conversation implementation. A host can mount the editor without
an agent and later connect its runtime directly or through standard MCP.

## Integration targets

| Target           | Requirements and current boundary                                                                                                                                                                       |
| ---------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Browser renderer | A browser supporting the ES2022 renderer bundle and its web dependencies; the host implements the services used by its application. Browsers do not acquire native file or print services from the SDK. |
| Electron host    | Injected host services and the same public mount API. The existing standalone desktop app is a separate consumer with its own preload and main-process requirements.                                    |
| Qt host          | Qt 6.8+ with Widgets, WebEngineWidgets, WebChannel, Network and PrintSupport; matching native toolchain and deployed Qt runtime. The reusable widget and sample host are included as sources.           |
| Node services    | Node 24+; documented peer dependencies for typed MCP transport integration and Node types for TypeScript consumers.                                                                                     |
| Other editors    | Sheets, Slides, PDF, Markdown and HTML remain inherited desktop implementations. No equivalent portable SDK or completion claim is made for them.                                                       |

These are integration requirements, not evidence that every OS/engine combination
has been tested. Use the source revision and exact host, engine and client versions
recorded with a delivery's GitHub acceptance evidence. An unavailable or skipped
target is not a verified target. The Qt sample currently opens `.docx`; legacy
Word, ODF, Hancom and HTML/MHT chunk conversion require explicit host integration.
The Electron application retains its own converters.

## Public editor lifecycle

Mount through `mountDocsEditor` with a host adapter, await the mount and
`editor.ready`, then use the public document, runtime and UI interfaces. Dispose
the editor when removing the view. One host/editor is supported per renderer
window; separate windows provide separate instances.

The SDK guide documents required file/native services and optional hosted
documents, resources, checkers, inspection, annotations, ribbon and file-menu
contributions. Optional capabilities must be provided explicitly. For managed
documents, the host owns opaque identity and complete DOCX bytes; unsupported
operations do not silently fall back to local storage.

Successful open/new completion publishes coherent document metadata and runtime
state. Captured scopes bind to a document generation and, when relevant, its
selection. They track valid edits without following a newly opened document.
Document replacement, authority withdrawal and disposal invalidate outstanding
authority. Callers must observe cancellation and freshness results rather than
blindly retrying against a different document.

## Choose an automation surface

| Surface                      | Data it operates on                                                                            | Integration contract                                                                                                 |
| ---------------------------- | ---------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------- |
| Office CLI and its MCP tools | File-level operations against the specified input; they do not see an editor's unsaved content | [CLI reference](../packages/cli/README.md); existing `genoffice` command names remain compatibility names            |
| Inherited desktop automation | The desktop application's configured file or visible-editor services                           | Desktop settings/help for that consumer; not an embedded SDK prerequisite                                            |
| Embedded Docs runtime        | The mounted live document, including unsaved edits and captured authorized selection           | Public runtime callers and scopes in the [SDK guide](../apps/docs/sdk/README.md#invoke-the-live-document)            |
| Embedded Docs MCP            | The authorized live runtime via standard stdio or host-controlled Streamable HTTP              | [SDK MCP reference](../apps/docs/sdk/README.md#standard-mcp), including supported protocol revisions and host grants |

The private host command bridge is local IPC, not a public agent protocol or the
HTTP MCP URL. Embedded HTTP MCP binds loopback on an allocated port and uses a
host-issued credential. Do not reuse the desktop CLI server's URL, upload API or
session rules. The SDK guide records the supported transport revisions; a new
revision or client is accepted only after actual interoperability verification.

Qt's trusted `startMcp`/`mcpState` result supplies separate editing and
`readOnlyConnection` configurations for the same captured live document. Each
uses its own credential and session callers; distribute only the intended
configuration to each client. The read-only server does not register write,
save or CLI tools. Both services end with their shared document binding, while
one client's reads do not refresh another client's baseline. This permission
grant is distinct from marking the hosted document itself read-only.

Hosts may forward editor-originated actions through `harness.submit` and may
resolve built-in instructions with stable operation identifiers. User-entered
instructions remain user input, not localized operation identifiers. Models,
templates, conversations and inference remain the embedding application's work.

## Permissions, persistence and privacy

File roots, exact dialog approvals and private-state exclusions are enforced by
the provided Node services. Hosts using their own services must enforce their
own authorization boundary. Resource selection does not grant arbitrary file
access; the bounded Office CLI is not a shell executor.

Qt hosts use `configureResources` to select already-authorized attachments and
knowledge files. Files outside read roots require a prior exact read grant;
selection supplies no write authority. Clear selections and revoke temporary
file grants when their application scope ends. Stopping MCP revokes its access,
not the host's independent resource configuration.

Save results distinguish committed, pending and unknown outcomes. Cancellation
does not roll back committed work. Preserve an outstanding draft's explicit
finalize/retain/discard decision before saving and obtain fresh context after a
stale-target result. See [Security](../SECURITY.md) and [Privacy](../PRIVACY.md)
for the trust boundaries and host/agent data responsibilities.

Versioning, package contents and delivery acceptance are described in
[SDK releases](docs-sdk-releases.md). The SDK is available through local build
and installation; a source checkout or package's existence is not a release or
an accepted complete-Docs milestone.
