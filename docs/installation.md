# Installing KĀRYA on Windows

KĀRYA is a modified fork of [GenOffice](../README.md) (Apache-2.0). This page
documents installing **KĀRYA's own Windows builds** — not the upstream
GenOffice releases listed on the README's download table.

## Prerequisites

- Windows 10 or later, x64.
- No Node.js, npm, Git, Python, or any other runtime is required — the
  installer ships a self-contained Electron app.
- **Ollama is optional but recommended.** KĀRYA never bundles Ollama or any
  AI model. If you want fully local AI, install Ollama separately (see
  [docs/ollama.md](ollama.md)). Cloud providers (Genspark, OpenAI, Claude,
  Gemini, DeepSeek) work without Ollama.

## Installer usage

1. Run `KĀRYA Setup <version>.exe` (NSIS installer).
2. Choose the installation directory (per-user install by default; you can
   change the destination during setup).
3. The installer creates a **KĀRYA** desktop shortcut and registers the file
   associations (`.docx`, `.xlsx`, `.pptx`, `.pdf`, `.md`, and related
   formats).

The app stores nothing in `Program Files`: all user configuration, recent
files, and AI settings live under `%APPDATA%\KĀRYA` (see "User data" below).

### First launch

- Sign in to a Genspark account **only if** you want the Genspark cloud AI
  and its search/image tools. This is optional — you can skip it and use
  Ollama, a direct vendor API key, or a custom endpoint instead.
- Open the AI settings (gear icon in the AI panel) to pick a provider.

### Uninstall

Uninstall through Windows **Settings → Apps → KĀRYA**. Uninstalling does not
delete your documents or your `%APPDATA%\KĀRYA` configuration.

## User data

| What                                   | Where                                     |
| -------------------------------------- | ----------------------------------------- |
| Settings, AI provider config, recents  | `%APPDATA%\KĀRYA`                         |
| Documents (default save location)      | `Documents\KĀRYA` (falls back to the old `Documents\GenOffice` if that is where your files already are) |

If you previously used GenOffice, KĀRYA migrates your configuration from
`%APPDATA%\GenOffice*` to the `KĀRYA` directories on first run — it copies
the data only when the new location is empty, so nothing is lost either way.

## Setting up local AI (Ollama)

1. Install [Ollama](https://ollama.com/download) for Windows and start it
   (the Ollama tray app runs the server automatically).
2. In any KĀRYA app, open the AI settings and select **Ollama** as the
   provider. The default endpoint is `http://localhost:11434/v1`.
3. Pull a model if you have none yet, e.g. `ollama pull llama3.2` (any
   model works — nothing is hard-coded).
4. Click **Refresh** in the model list, choose a model, and chat.

See [docs/ollama.md](ollama.md) for details and common errors.

## Cloud provider configuration

- **Genspark** — sign in with a Genspark account; no API key to enter.
- **OpenAI / Claude / Gemini / DeepSeek** — select the provider, paste your
  API key, pick a model.
- **Custom** — set a base URL of any OpenAI-compatible server and (optionally)
  an API key.

API keys are stored only in the local settings file under your user profile;
they are never sent anywhere except the provider you configured. There is no
automatic fallback between providers.
