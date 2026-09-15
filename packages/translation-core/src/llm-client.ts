/**
 * LLM client seam for `@genoffice/translation-core`.
 *
 * Why this file exists (W9 deliverable — migration to `pi-ai`):
 *   The translation core historically delegates every chat call to
 *   `@genoffice/ai-provider`'s `chatForProvider`. As part of the broader
 *   migration to `@earendil-works/pi-ai` (the LLM SDK shipped with the pi
 *   coding agent), this module is the seam where that swap happens.
 *
 *   Today: `aiProviderCaller` is the default and only production caller.
 *   Future: `piAiCaller` will become the default once the provider mapping
 *   (anthropic / gemini / openai / …) is complete. Until then, callers can
 *   opt into `piAiCaller` per-request via `callLlmWith()` for early adopter
 *   code paths.
 *
 * The public shape — `LlmCallOptions` / `LlmCallResult` — is the stable
 * contract between translation-core and the underlying SDK. Tests mock
 * `../src/llm-client` rather than `@genoffice/ai-provider` so the migration
 * doesn't require touching every test.
 */

import {
  chatForProvider,
  isAiOverloadedError,
  type AiProviderConfig,
  type AiProviderId,
} from '@genoffice/ai-provider'

// ---------------------------------------------------------------------------
// Public shape — stable across SDK migrations
// ---------------------------------------------------------------------------

export interface LlmCallOptions {
  provider: AiProviderId
  config: AiProviderConfig
  systemPrompt: string
  userPrompt: string
  /**
   * Optional metadata that flows through to the underlying provider. Hosts
   * use this to thread tags like `glossaryCategory` or `qualityCheck` so
   * they show up in provider-side telemetry without changing the call
   * shape.
   */
  metadata?: Record<string, string>
}

export interface LlmCallResult {
  ok: boolean
  /** Final assistant text when ok=true. */
  content?: string
  /** Human-readable error when ok=false. */
  error?: string
  /** True when the underlying provider indicated overloaded / rate-limited. */
  overloaded?: boolean
}

/** Pluggable LLM caller. */
export type LlmCaller = (opts: LlmCallOptions) => Promise<LlmCallResult>

// ---------------------------------------------------------------------------
// aiProviderCaller — current production implementation
// ---------------------------------------------------------------------------

/**
 * Default caller: delegates to `@genoffice/ai-provider`'s `chatForProvider`.
 * Kept in place until `piAiCaller` reaches feature parity.
 */
export const aiProviderCaller: LlmCaller = async (opts) => {
  try {
    const result = await chatForProvider(
      opts.provider,
      opts.config,
      opts.systemPrompt,
      opts.userPrompt,
    )
    if (!result.ok) {
      return {
        ok: false,
        error: typeof result.error === 'string' ? result.error : 'Translation failed',
        overloaded: isAiOverloadedError(result.error),
      }
    }
    return { ok: true, content: result.content ?? '' }
  } catch (err) {
    return {
      ok: false,
      error: err instanceof Error ? err.message : String(err),
      overloaded: isAiOverloadedError(err),
    }
  }
}

// ---------------------------------------------------------------------------
// piAiCaller — future production implementation (W9.5+)
// ---------------------------------------------------------------------------

/**
 * Future caller: delegates to `@earendil-works/pi-ai`'s `completeSimple`.
 * Currently throws because the GenOffice → pi provider mapping (17 providers)
 * is not yet complete. Once it is, this becomes the default.
 */
export const piAiCaller: LlmCaller = async (_opts) => {
  throw new Error(
    '[translation-core] piAiCaller is not yet implemented. ' +
      'The migration from @genoffice/ai-provider to @earendil-works/pi-ai is tracked in agent1.md §16.11.',
  )
}

// ---------------------------------------------------------------------------
// Active caller — swappable via setLlmCaller()
// ---------------------------------------------------------------------------

let activeCaller: LlmCaller = aiProviderCaller

/** Replace the active LLM caller (used by tests + the future migration). */
export function setLlmCaller(caller: LlmCaller): void {
  activeCaller = caller
}

/** Return the currently active LLM caller. */
export function getLlmCaller(): LlmCaller {
  return activeCaller
}

/**
 * Public entry point used by `provider.ts`. Routes to the active caller.
 * Tests can stub this entire module to assert on the call args.
 */
export async function callLlm(opts: LlmCallOptions): Promise<LlmCallResult> {
  return activeCaller(opts)
}

/**
 * One-shot caller — bypasses the global active caller. Useful for
 * per-request opt-in (e.g. a beta app that wants pi-ai for one provider
 * only).
 */
export async function callLlmWith(caller: LlmCaller, opts: LlmCallOptions): Promise<LlmCallResult> {
  return caller(opts)
}
