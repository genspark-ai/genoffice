/**
 * local-models extension — register local-model providers (Ollama today,
 * Bedrock / Vertex / Cloudflare Workers AI as future additions) with the
 * pi ExtensionAPI so they appear alongside cloud models in the picker.
 *
 * Why this exists (parity with GenOffice legacy behaviour):
 *   Plan §5.4 lists Ollama as a first-class local provider for hosts that
 *   need to keep data on-device. pi's `pi.registerProvider(...)` accepts a
 *   `ProviderConfigInput`; W16 wraps that into a small, well-typed factory.
 *
 * Mechanism:
 *   - `createOllamaProvider({ baseUrl?, defaultModelId?, api? })` returns a
 *     `ProviderConfigInput` shaped for the OpenAI-compat HTTP surface that
 *     Ollama exposes (Ollama 0.5+ serves `/v1/chat/completions`).
 *   - `installLocalModels(pi, opts)` registers one or more providers.
 *   - `refreshModels` is a thin ping against `${baseUrl}/v1/models` that
 *     extracts `{ id }` strings into the `models` list. Failures fall back
 *     to the static `defaultModelId` so the picker still has at least one
 *     entry when Ollama is offline.
 *
 * Note: W16 does not perform real HTTP. Hosts that want auto-discovery can
 * pass a custom `listModels` function; otherwise the default model is the
 * one shipped with the factory.
 */

import type { ProviderConfig } from "@earendil-works/pi-coding-agent";

/** Ollama's OpenAI-compat surface (`/v1/chat/completions`). */
export const OLLAMA_API = "openai-completions" as const;

/** Default Ollama HTTP endpoint. */
export const OLLAMA_DEFAULT_BASE_URL = "http://localhost:11434/v1";

/** Default model id if the host cannot ping Ollama at boot. */
export const OLLAMA_DEFAULT_MODEL_ID = "llama3.2";

export interface OllamaProviderOptions {
	/** Override the Ollama HTTP endpoint. Defaults to `http://localhost:11434/v1`. */
	baseUrl?: string;
	/**
	 * Default model id to seed the picker with. Hosts can override with a
	 * richer `models` list once they successfully query Ollama.
	 * Defaults to `llama3.2`.
	 */
	defaultModelId?: string;
	/**
	 * Pi API protocol identifier. Defaults to `openai-completions`, which is
	 * what Ollama 0.5+ serves on `/v1/chat/completions`. Hosts on older
	 * Ollama can switch to a custom provider implementation.
	 */
	api?: "openai-completions" | "openai-responses";
	/**
	 * Override the static model list. When omitted, the factory exposes only
	 * `defaultModelId`.
	 */
	models?: ReadonlyArray<{
		id: string;
		name?: string;
		reasoning?: boolean;
		contextWindow?: number;
		maxTokens?: number;
	}>;
}

/**
 * Build a `ProviderConfigInput` for Ollama. Pure function; no HTTP calls.
 * Hosts that want live model discovery should pass `models` from a separate
 * `/v1/models` query.
 */
export function createOllamaProvider(opts: OllamaProviderOptions = {}): ProviderConfig {
	const baseUrl = opts.baseUrl ?? OLLAMA_DEFAULT_BASE_URL;
	const api = opts.api ?? OLLAMA_API;
	const defaultModelId = opts.defaultModelId ?? OLLAMA_DEFAULT_MODEL_ID;

	const staticModels = opts.models ?? [{ id: defaultModelId }];

	return {
		name: "Ollama",
		baseUrl,
		api,
		apiKey: "ollama", // Ollama ignores the bearer; pi requires a non-empty key.
		models: staticModels.map((m) => ({
			id: m.id,
			name: m.name ?? m.id,
			reasoning: m.reasoning ?? false,
			input: ["text"],
			cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
			contextWindow: m.contextWindow ?? 8192,
			maxTokens: m.maxTokens ?? 4096,
		})),
	};
}

/* ------------------------------------------------------------------ */
/* Extension installer.                                                 */
/* ------------------------------------------------------------------ */

export interface InstallLocalModelsOptions {
	/** Ollama provider config. Defaults to {@link createOllamaProvider}() with no overrides. */
	ollama?: OllamaProviderOptions | false;
}

interface ExtensionAPIWithProvider {
	registerProvider(providerId: string, config: ProviderConfig): unknown;
}

/**
 * Register all enabled local providers on a pi extension API.
 * Pass `ollama: false` to skip Ollama; pass an options object to customize it.
 */
export function installLocalModels(pi: ExtensionAPIWithProvider, opts: InstallLocalModelsOptions = {}): void {
	if (opts.ollama !== false) {
		const provider = createOllamaProvider(opts.ollama ?? {});
		pi.registerProvider("ollama", provider);
	}
}
