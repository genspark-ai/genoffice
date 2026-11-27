# AI models and settings

## Providers and models

All AI model providers, custom endpoints, and API keys are configured inside the **Settings** dialog (opened from the account row at the bottom left of the Home screen):

![The Settings window](img/settings-general.png)

- **Genspark hosted models**: For a zero-configuration setup, sign in with your Genspark account using the browser device-code flow and start using the assistant right away—no API keys or endpoint URLs to manage.
- **Custom endpoints (BYOK)**: If you prefer to use your own API keys or an internal company gateway, open **Settings ▸ AI Model** (and **AI Media & Search**) to enter a base URL and API key per protocol—including OpenAI-compatible endpoints, Anthropic, Google Gemini, DeepSeek, DashScope (Qwen), MiniMax, Atlas Cloud, and more. Your keys are stored exclusively in GenOffice's owner-only settings file on your local machine and are sent only in the request headers addressed to the endpoint you configured.
- **Mix and match models per capability**: You can assign a different provider and model to each distinct job—picking one model for document chat and drafting, a dedicated image model for **image generation**, and a vision-capable model for **image and video analysis**.
- **MiniMax H3-Image and MiniMax H3-Image Highspeed**: Alongside `image-01`, the MiniMax image provider now supports **MiniMax H3-Image** (the default high-detail model) and **MiniMax H3-Image Highspeed** for faster turnaround on both international and CN (`api.minimax.cn`) endpoints; both H3 models request crisp **2K-area output** by default and accept up to **16 reference images** when performing an image-to-image edit.
- **One-click connection test**: Before saving a new provider configuration, click **Connection test** to verify on the spot that your base URL is reachable, your API key is accepted, and the chosen model responds properly.
- **Gateway-friendly base URLs**: Custom base URLs can include subpaths and query strings (such as corporate proxy routes or API gateways); GenOffice preserves your path prefix and query parameters cleanly when appending endpoint routes.

## CLI integration (Codex-class)

- **Pointing to a local CLI binary**: You can also specify the file path to a locally installed Codex-class CLI executable in Settings (paths inside non-ASCII home directories and paths starting with `~` both work out of the box, with `~` expanded automatically); clicking **Detect models** probes the executable and populates its available models for you.
- **Existence-based path validation**: Path validation checks only whether the binary exists on disk, without imposing artificial ASCII-only character restrictions on your folder names.

## When changes apply

- **Immediate cutover on the next turn**: Switching models, endpoints, or API keys in Settings takes effect immediately across every open editor; if an AI conversation is already streaming a reply when you save, that in-flight turn finishes cleanly on the previous configuration and your very next prompt uses the new model.
