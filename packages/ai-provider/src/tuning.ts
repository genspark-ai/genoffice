/**
 * Per-request generation knobs (temperature / output ceiling / reasoning
 * effort), resolved the same way by the one-shot and streaming paths.
 *
 * The awkward part these exist for: reasoning models (GPT-5, o-series, and the
 * OpenAI-compatible endpoints that imitate them) reject `temperature` outright
 * unless it is their own default, so a fixed value cannot be sent to every
 * backend. `null` therefore has to mean "leave the field out", distinct from
 * `undefined` meaning "caller expressed no preference, use the house default".
 */
import type { AiProviderConfig, ReasoningEffort } from './types'

/** house default for chat-style models, kept from before these knobs existed */
export const DEFAULT_TEMPERATURE = 0.3

const REASONING_EFFORTS: readonly ReasoningEffort[] = ['minimal', 'low', 'medium', 'high']

export function isReasoningEffort(value: unknown): value is ReasoningEffort {
  return REASONING_EFFORTS.includes(value as ReasoningEffort)
}

/**
 * `{ temperature }` to spread into a request body, or `{}` when the field must
 * not be sent at all.
 */
export function temperatureField(config: AiProviderConfig): { temperature?: number } {
  if (config.temperature === null) return {}
  return { temperature: config.temperature ?? DEFAULT_TEMPERATURE }
}

/** the configured output ceiling, falling back to what the caller asked for */
export function resolveMaxTokens(config: AiProviderConfig, requested: number): number {
  const configured = config.maxTokens
  return typeof configured === 'number' && configured > 0 ? configured : requested
}

/**
 * `{ max_tokens }` for request shapes that have no built-in default of their
 * own (the one-shot OpenAI-compatible call): sent only when the user set a
 * ceiling, so an unconfigured endpoint keeps letting the server decide.
 */
export function maxTokensField(config: AiProviderConfig): { max_tokens?: number } {
  const configured = config.maxTokens
  return typeof configured === 'number' && configured > 0 ? { max_tokens: configured } : {}
}

/** `{ reasoning_effort }` for OpenAI-compatible bodies, or `{}` when unset */
export function reasoningEffortField(config: AiProviderConfig): {
  reasoning_effort?: ReasoningEffort
} {
  return config.reasoningEffort ? { reasoning_effort: config.reasoningEffort } : {}
}
