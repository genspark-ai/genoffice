/**
 * Tests for the local-models extension (W16 deliverable).
 *
 * Verifies:
 *   - createOllamaProvider returns a well-shaped ProviderConfigInput
 *   - Defaults match plan §5.4 (OpenAI-compat API on localhost:11434/v1)
 *   - Custom baseUrl / defaultModelId / api / models thread through
 *   - installLocalModels registers exactly one provider by default
 *   - installLocalModels respects the `ollama: false` opt-out
 */

import { describe, expect, it, vi } from "vitest";
import type { ProviderConfig } from "@earendil-works/pi-coding-agent";
import {
	createOllamaProvider,
	installLocalModels,
	OLLAMA_API,
	OLLAMA_DEFAULT_BASE_URL,
	OLLAMA_DEFAULT_MODEL_ID,
} from "../src/extensions/local-models";

interface Harness {
	registerProvider: ReturnType<typeof vi.fn>;
}

function makeHarness(): Harness {
	const registerProvider = vi.fn();
	return { registerProvider };
}

describe("module constants", () => {
	it("OLLAMA_DEFAULT_BASE_URL points at localhost:11434/v1", () => {
		expect(OLLAMA_DEFAULT_BASE_URL).toBe("http://localhost:11434/v1");
	});

	it("OLLAMA_API is openai-completions", () => {
		expect(OLLAMA_API).toBe("openai-completions");
	});

	it("OLLAMA_DEFAULT_MODEL_ID is llama3.2", () => {
		expect(OLLAMA_DEFAULT_MODEL_ID).toBe("llama3.2");
	});
});

describe("createOllamaProvider", () => {
	it("returns a provider named 'Ollama' with the default endpoint and model", () => {
		const provider = createOllamaProvider();
		expect(provider.name).toBe("Ollama");
		expect(provider.baseUrl).toBe(OLLAMA_DEFAULT_BASE_URL);
		expect(provider.api).toBe(OLLAMA_API);
		expect(provider.apiKey).toBeTruthy();
		expect(provider.models).toHaveLength(1);
		expect(provider.models?.[0]?.id).toBe(OLLAMA_DEFAULT_MODEL_ID);
	});

	it("threads baseUrl through to the provider config", () => {
		const provider = createOllamaProvider({ baseUrl: "http://gpu-box.lan:11434/v1" });
		expect(provider.baseUrl).toBe("http://gpu-box.lan:11434/v1");
	});

	it("threads defaultModelId into the static models list when no models override is supplied", () => {
		const provider = createOllamaProvider({ defaultModelId: "qwen2.5-coder:32b" });
		expect(provider.models).toHaveLength(1);
		expect(provider.models?.[0]?.id).toBe("qwen2.5-coder:32b");
		expect(provider.models?.[0]?.name).toBe("qwen2.5-coder:32b");
	});

	it("threads a custom api when supplied", () => {
		const provider = createOllamaProvider({ api: "openai-responses" });
		expect(provider.api).toBe("openai-responses");
	});

	it("accepts a custom models list with context windows and reasoning flags", () => {
		const provider = createOllamaProvider({
			models: [
				{ id: "llama3.2", reasoning: false, contextWindow: 8192, maxTokens: 4096 },
				{ id: "deepseek-r1:14b", name: "DeepSeek R1 14B", reasoning: true, contextWindow: 65536, maxTokens: 8192 },
			],
		});
		expect(provider.models).toHaveLength(2);
		const ds = provider.models?.[1];
		expect(ds?.id).toBe("deepseek-r1:14b");
		expect(ds?.name).toBe("DeepSeek R1 14B");
		expect(ds?.reasoning).toBe(true);
		expect(ds?.contextWindow).toBe(65536);
		expect(ds?.maxTokens).toBe(8192);
		expect(ds?.input).toEqual(["text"]);
		expect(ds?.cost).toEqual({ input: 0, output: 0, cacheRead: 0, cacheWrite: 0 });
	});

	it("uses the supplied model id as the display name when no name is given", () => {
		const provider = createOllamaProvider({
			models: [{ id: "phi3:mini" }],
		});
		expect(provider.models?.[0]?.name).toBe("phi3:mini");
	});

	it("falls back to default contextWindow / maxTokens when caller omits them", () => {
		const provider = createOllamaProvider({
			models: [{ id: "tiny" }],
		});
		expect(provider.models?.[0]?.contextWindow).toBe(8192);
		expect(provider.models?.[0]?.maxTokens).toBe(4096);
	});

	it("always returns a non-empty apiKey (Ollama ignores bearer but pi requires it)", () => {
		const provider = createOllamaProvider();
		expect(typeof provider.apiKey).toBe("string");
		expect((provider.apiKey ?? "").length).toBeGreaterThan(0);
	});
});

describe("installLocalModels", () => {
	it("registers an ollama provider by default", () => {
		const { registerProvider } = makeHarness();
		const pi = { registerProvider } as unknown as { registerProvider: (id: string, cfg: ProviderConfig) => unknown };
		installLocalModels(pi);
		expect(registerProvider).toHaveBeenCalledTimes(1);
		expect(registerProvider.mock.calls[0]?.[0]).toBe("ollama");
		const config = registerProvider.mock.calls[0]?.[1] as ProviderConfig;
		expect(config.baseUrl).toBe(OLLAMA_DEFAULT_BASE_URL);
	});

	it("threads ollama options into the registered provider", () => {
		const { registerProvider } = makeHarness();
		installLocalModels({ registerProvider } as unknown as { registerProvider: (id: string, cfg: ProviderConfig) => unknown }, { ollama: { baseUrl: "http://remote:9999/v1" } });
		const config = registerProvider.mock.calls[0]?.[1] as ProviderConfig;
		expect(config.baseUrl).toBe("http://remote:9999/v1");
	});

	it("skips registration when ollama: false", () => {
		const { registerProvider } = makeHarness();
		installLocalModels({ registerProvider } as unknown as { registerProvider: (id: string, cfg: ProviderConfig) => unknown }, { ollama: false });
		expect(registerProvider).not.toHaveBeenCalled();
	});
});
