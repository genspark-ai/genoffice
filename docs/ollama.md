# Local AI with Ollama

KĀRYA can run its entire AI panel against a locally installed
[Ollama](https://ollama.com) server. This page covers the endpoint, model
discovery, provider selection, and common errors.

## The endpoint

Ollama's OpenAI-compatible API is served at:

```
http://localhost:11434/v1
```

This is the default `baseUrl` when you select **Ollama** as the provider.
KĀRYA discovers installed models by querying the server's native endpoint
(`http://localhost:11434/api/tags`). You can point `baseUrl` at a different
host (for example a remote Ollama server on your network); the discovery
endpoint is derived from it.

## Installing and starting Ollama

1. Download and install Ollama from <https://ollama.com/download> (Windows /
   macOS / Linux builds all work with KĀRYA).
2. Start it. On Windows the Ollama tray app starts the server automatically;
   from a terminal you can also run `ollama serve`.
3. Pull a model if you have none, for example:

   ```bash
   ollama pull llama3.2
   ```

   Any model works — KĀRYA never assumes a specific model is installed.

## Selecting Ollama in KĀRYA

1. Open the AI panel in any app and click the settings (gear) icon.
2. In **Provider**, choose **Ollama**.
3. The dialog shows a **connection status** chip:
   - **Connected** — the server answered and models were listed.
   - **Ollama is not running** — the server could not be reached (connection
     refused / timed out). Start Ollama and try again.
   - **Invalid endpoint** — the server answered, but the endpoint is wrong
     (bad base URL, wrong port, or an HTTP error).
4. The **model list** is populated from the server (model name + parameter
   size). If the list is empty, click **Refresh**. If your saved model is not
   installed, KĀRYA shows a notice and keeps your selection — it never
   silently switches to another model or to a cloud provider.
5. Select a model and start chatting. No API key is required for a local
   server.

The model list is cached briefly (10 seconds) so opening the dialog is fast;
**Refresh** always probes the server again. KĀRYA does not poll Ollama in the
background, never auto-downloads models, never removes models, and never
auto-starts Ollama.

## Capabilities

KĀRYA does not guess at model capabilities. The model list shows only the
name and parameter size reported by the server — no chat/streaming/tools/
vision badges, because those cannot be detected reliably without claiming
unsupported behavior.

## Common errors

| Symptom                          | Cause                                        | Fix                                          |
| -------------------------------- | -------------------------------------------- | -------------------------------------------- |
| "Ollama is not running"          | Server not started, wrong host, or wrong port | Start Ollama; check `baseUrl` in the AI settings |
| "Invalid endpoint"               | Base URL points at a non-Ollama path         | Reset to `http://localhost:11434/v1`         |
| Model list empty                 | No models pulled yet                         | `ollama pull <model>` then click Refresh     |
| "The selected model is not installed" | Saved model was removed from the server  | Pull it again, or pick another model         |
| Slow first response              | Model needs to load into memory              | Normal; subsequent turns are faster          |

## Privacy

With the **Ollama** provider, every request goes to your local (or
user-configured) server. Your prompts and document content never leave your
machine, and there is **no automatic cloud fallback** — if Ollama is down,
KĀRYA tells you, it does not silently route to a cloud provider.
