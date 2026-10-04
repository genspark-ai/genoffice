# Settings, language, theme and MCP integrations

## Opening settings

The gear button on Home opens the settings panel; AI-related options live in its AI section

![The Settings window](img/settings-integrations.en.png) — model configuration is covered in AI models and settings.

## Language

- The settings offer **21 UI languages**: English, Simplified Chinese, Japanese, Korean, French, German, Spanish, Thai, Indonesian, Russian, Arabic, Portuguese, Italian, Polish, Czech, Dutch, Malay, Hebrew, Hindi, Traditional Chinese, Vietnamese.
- Switching applies immediately and persists; the native menu bar rebuilds with the language.
- This manual currently ships Simplified Chinese and English bodies; every other locale shows English.

## Theme

Light / Dark / System. System follows the OS appearance, and the editors re-skin in sync without flashing.

## Default app bindings

Settings can register GenOffice as the handler for .docx / .xlsx / .pptx / .pdf and friends (platform-level default-app registration; confirm when prompted).

## Third-party notices and updates

- Help ▸ Third-party notices: the full OSS license inventory shipped with the app.
- Help ▸ Check for updates: triggers a manual check; a newer version prompts to install.

## Genspark sign-in

- The sign-in entry (settings or the cloud project list) uses a **device-code** flow: GenOffice shows a code and opens the browser login; it continues automatically once done.
- Sign-in is used only for: the cloud project list and Genspark's hosted models. Without it, every local feature and custom models keep working.
- Sign out is one click in settings.

## MCP integration (for advanced users / AI clients)

GenOffice embeds a local **MCP server** so external AI clients (Claude Desktop, Cursor, ...) can read and write your documents directly:

- Start: `genoffice mcp` on the command line (port and auth token configurable; loopback-only by default).
- Capabilities: create/open/edit docx, xlsx and pptx, read contents, convert formats, export PDF and more — the same toolset the desktop apps use.
- Security: token auth is optional but recommended; the listener stays on the local machine by default; see `genoffice mcp --help`.

## Command-line cheat sheet

| Command            | What it does               |
| ------------------ | -------------------------- |
| `genoffice <file>` | open a file                |
| `genoffice mcp`    | start the local MCP server |
| `genoffice --help` | every command and flag     |
