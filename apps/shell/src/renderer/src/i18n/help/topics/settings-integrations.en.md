# Settings, language, theme and MCP integrations

## Opening settings

The account row at the bottom left of Home opens the settings panel (it reads Sign in when you are logged out). It has six sections: Account, AI Model, AI Media & Search, General, Integrations and About.

![Settings ▸ General, where language, theme, AutoSave and the usage-statistics switch live](img/settings-general.png)

AI model configuration has its own article; **AI Media & Search** is where you turn image generation, image analysis, video analysis, web search and local file search on per provider.

## Language

- The settings offer **21 UI languages**: English, Simplified Chinese, Japanese, Korean, French, German, Spanish, Thai, Indonesian, Russian, Arabic, Portuguese, Italian, Polish, Czech, Dutch, Malay, Hebrew, Hindi, Traditional Chinese, Vietnamese.
- Switching applies immediately and persists; the native menu bar rebuilds with the language.

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

**Integrations** is the pane that connects GenOffice to a coding agent, and it has its own article: Connecting a coding agent. The short version — pick a route (the command line, or MCP), follow that section, then start a new chat and ask.

![Settings ▸ Integrations: the three steps, then the skill rows and the MCP options](img/settings-integrations.png)

Under **Local HTTP server** the app can run the server itself — an enable switch, a port — and **Advanced** adds the health-check URL and the log file. It listens on localhost only.

## Command-line cheat sheet

| Command            | What it does               |
| ------------------ | -------------------------- |
| `genoffice <file>` | open a file                |
| `genoffice mcp`    | start the local MCP server |
| `genoffice --help` | every command and flag     |
