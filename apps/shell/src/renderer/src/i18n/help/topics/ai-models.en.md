# AI models and settings

## Providers and models

Models and keys are configured in Settings (the account row at the bottom left of Home):

![The Settings window](img/settings-general.png)

- **Genspark hosted**: sign in (device-code flow) and use it — zero configuration.
- **Custom endpoints (BYOK)**: Settings ▸ AI takes a base URL and API key per protocol — OpenAI-compatible, Anthropic, Gemini, DeepSeek, DashScope (qwen) and more. Keys live only in request headers — never on disk, in logs or child-process env.
- A different model can be picked per capability: chat/generation, image generation, image analysis.
- **Connection test**: verify endpoint reachability and model visibility before saving.
- Base URLs may carry a path and query string (gateway-style); endpoint paths are appended correctly.

## CLI integration (Codex-class)

- Settings accept a local CLI program path (non-ASCII home directories and a ~ prefix work; ~ is expanded automatically); Detect models probes the CLI's available models.
- Validation checks existence only — no charset restrictions.

## When changes apply

- Model and endpoint changes apply immediately; an in-flight conversation keeps the old configuration until its next turn.
