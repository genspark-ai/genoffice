# AI providers

Every app's AI panel is provider-agnostic: the UI talks to a single provider
abstraction (`packages/ai-provider`), and the selected provider handles all
model calls. The agent core never branches on provider ids.

## Supported providers

| Provider  | Authentication                     | Model list                                   |
| --------- | ---------------------------------- | -------------------------------------------- |
| Genspark  | Genspark account (device-code sign-in; no key to manage) | Claude / GPT / Gemini families through the Genspark proxy |
| Claude    | Anthropic API key (`sk-ant-...`)   | Claude models                                |
| Gemini    | Google API key (`AIza...`)         | Gemini models                                |
| DeepSeek  | DeepSeek API key (`sk-...`)        | `deepseek-chat`, `deepseek-reasoner`         |
| OpenAI    | OpenAI API key (`sk-...`)          | GPT-4.x models                               |
| Custom    | Optional API key                   | User-entered model name; any OpenAI-compatible server (configurable base URL) |
| Ollama    | None (local server)                | Discovered live from the server (`/api/tags`) |

## Configuration model

Each provider has a `baseUrl` (where needed), an `apiKey`, and a `model`:

- **Ollama** — `baseUrl` configurable (default `http://localhost:11434/v1`),
  `apiKey` optional, `model` discovered from the server.
- **Cloud providers** — existing authentication is preserved; Genspark's key
  is injected by the main process from the account login (never stored in
  settings), direct vendors store the key you paste in the local settings
  file.
- **Custom** — user-configured endpoint and model name.

Settings are stored per app in the user-data directory (e.g.
`%APPDATA%\KĀRYA` on Windows). Saved settings from before the multi-provider
model are migrated automatically: the legacy single-endpoint configuration
maps into the **Custom** provider slot.

## Privacy

- **Ollama** selected → all inference happens on your local (or
  user-configured) server. Prompts and document content are not sent to any
  cloud provider. There is **no automatic cloud fallback**.
- **Cloud provider** selected → prompts are sent to that provider according
  to its API, using the credentials you configured. Genspark model calls are
  attributed to GenOffice usage through an `X-Agent-Type` header sent only
  to the Genspark proxy.
- KĀRYA does not add telemetry for prompts or document content, and never
  sends document content to a cloud provider merely to test connectivity.

## How the selection works

The selected provider (`settings.provider`) routes every AI request — chat,
streaming, and tool calls. If a provider is not fully configured (missing
key or model), the UI marks the AI panel as not configured; nothing is sent.
If the selected Ollama model disappears from the server, you get a clear
notice and the previous selection stays — KĀRYA never silently switches
providers or models.
