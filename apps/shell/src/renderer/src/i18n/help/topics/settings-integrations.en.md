# Settings, language, theme and MCP integrations

## Opening settings

Clicking the account row at the bottom left of the Home screen opens the settings dialog (when you are signed out, this row reads **Sign in**). The dialog is organized into six dedicated sections along the left: **Account**, **AI Model**, **AI Media & Search**, **General**, **Integrations**, and **About**.

![Settings ▸ General, where language, theme, AutoSave and the usage-statistics switch live](img/settings-general.png)

Configuring chat and generation models has its own dedicated chapter; **AI Media & Search** is where you configure the providers and keys for image generation, image analysis, video analysis, web search, and local file search.

## Language

- GenOffice ships with **23 fully localized UI languages**: English, Simplified Chinese, Traditional Chinese, Japanese, Korean, French, German, Spanish, Thai, Indonesian, Russian, Arabic, Portuguese, Italian, Polish, Czech, Dutch, Malay, Hebrew, Hindi, Vietnamese, Turkish, and Hungarian.
- Picking a different language switches the interface immediately and saves your preference across launches; the native system menu bar, context menus, and this in-app manual all rebuild in the chosen language on the spot.

## Theme

Choose between **Light**, **Dark**, and **System** appearance modes. When set to **System**, GenOffice tracks your operating system's day/night schedule automatically, re-skinning every open editor toolbar, sidebar, and dialog in sync without a white flash—while keeping your actual document pages, slide canvases, and print outputs faithful to their authored colors.

## General

- **Send anonymous usage statistics**: Enabled by default, this uses Google Analytics 4 to report high-level application events (such as `"opened a .docx"`) along with standard network transport metadata and your public IP address; your document text, spreadsheet numbers, slide contents, and file names are never collected or transmitted, and you can switch telemetry off here at any time.
- **AI assistant layout preferences**: Customize your **AI sidebar position** (docked on the left or right edge of the editor), adjust the **AI panel text size** for comfortable reading, and toggle **Spell check in AI chat** for your prompt input box.
- **Open the AI panel in new documents**: Off by default so newly created documents start with a clean, full-width editing canvas and the assistant tucked one click away on the ribbon; turn it on if you prefer the AI panel to open automatically whenever you create a new file.
- **Auto-save all documents**: Enables AutoSave by default across every editor once a file has been saved manually for the first time, while still letting you toggle AutoSave off individually on any single editor window's toolbar.
- **Default save location and file associations**: Use the **Save Location** row and its **Change** button to pick the default folder offered when saving new documents, or click **Default app for Office documents** to register GenOffice as your system handler for `.docx`, `.xlsx`, and `.pptx` files.

## AI Media & Search

Rather than simple on/off toggles, each multimodal or search capability lets you pick the specific provider and model that serves it best, while sharing a provider's API key and base URL once across every row that uses that vendor:

- **Web search, Image generation, Image analysis, and Video analysis**: Each row has its own provider selector, model picker, API key field, and base URL so you can mix and match vendors freely—for example, pairing one provider's chat model with another vendor's dedicated image generator.
- **Local file search and Jev search reranking**: Full-text local file search runs entirely on your computer, and **Jev search reranking** underneath it is **off by default** so nothing ever leaves your machine unless you ask for it. When you switch reranking on, short excerpts from the top 20 local matches (up to 1,200 characters per document, along with their file and folder names) are sent to the decision endpoint you choose—Jev via OpenRouter or TypeSafe, Perplexity, Liquid AI d1, Cloudflare, a local engine, or a custom URL—to reorder the hits by semantic relevance.

## About

- **Version and open-source repository**: Displays the exact build version you are running alongside a direct link to the project's GitHub repository and a **Star on GitHub** button.
- **Update Channel**: Switch between the **Stable** and **Beta** release channels; changing the channel takes effect immediately and triggers an update check right away, though switching from Beta back to Stable will not forcibly downgrade an already-installed Beta build.

## Default app bindings

Settings can register GenOffice at the operating-system level as the default application for opening `.docx`, `.xlsx`, `.pptx`, `.pdf`, and related document formats; simply confirm the system prompt shown by macOS, Windows, or your Linux desktop environment when asked.

## Third-party notices and updates

- **Help ▸ Third-party notices**: Opens the complete inventory of open-source libraries and license texts bundled with GenOffice.
- **Help ▸ Check for updates**: Manually checks the active release channel for a newer build and prompts you to download and install it when one is available.

## Genspark sign-in

- **Browser device-code sign-in**: Signing in from either the Settings dialog or the cloud project list on Home uses a secure **device-code** flow—GenOffice displays a one-time verification code and opens the sign-in page in your default browser, then resumes automatically as soon as you approve the session.
- **What sign-in is used for**: A Genspark account is only needed for syncing your web **cloud project list** and calling **Genspark's hosted AI models**; without signing in, every local editing feature, format converter, and custom BYOK model endpoint continues to work in full.
- **Confirmed sign-out**: Clicking **Sign out** in Settings now opens a confirmation dialog first so a stray click never disconnects your workspace mid-task; once confirmed, it signs out every Genspark-backed AI feature at once (both cloud projects and hosted models) while leaving all local features and custom endpoints untouched.

## MCP integration (for advanced users / AI clients)

**Settings ▸ Integrations** is the control center for connecting GenOffice's document engines to external coding agents such as Claude Code, Cursor, Codex, and OpenCode (covered in full in the _Connecting a coding agent_ chapter). In short, you pick either the CLI skill route or the MCP route, click install for your agent, and start asking your agent to create, edit, or convert Office documents in plain language.

![Settings ▸ Integrations: the three steps, then the skill rows and the MCP options](img/settings-integrations.png)

Under **Local HTTP server**, GenOffice can also host the MCP server directly inside the desktop app—giving you an enable switch, a configurable port, and an **Advanced** section with the health-check URL and local activity log file—listening strictly on `localhost`.

## Command-line cheat sheet

| Command            | What it does               |
| ------------------ | -------------------------- |
| `genoffice <file>` | open a file                |
| `genoffice mcp`    | start the local MCP server |
| `genoffice --help` | every command and flag     |
