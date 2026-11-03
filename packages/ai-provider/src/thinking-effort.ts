/**
 * Per-model thinking-effort resolution: what the user picked in settings,
 * translated into the one request field each model family actually accepts.
 *
 * The levels follow OpenAI's current ladder (the widest one): `none` turns
 * reasoning off, `xhigh`/`max` are the upper tiers only the first-party
 * OpenAI Responses API accepts. Every other family maps onto that ladder with
 * its own shape, and a family without any effort parameter resolves to nothing
 * sent, i.e. today's behaviour.
 *
 * Only surfaces documented by the vendor's own API are listed. The vendors
 * whose effort field is still documentation-only (GLM, Doubao, MiniMax,
 * DeepSeek, Kimi, Hunyuan, Mistral) stay knobless until a live probe confirms
 * them: a wrong row costs a 400 on every request for that provider.
 *
 * Resolution order mirrors the temperature rule: a user's stored pick wins,
 * an absent pick means the vendor default, and nothing here overrides a
 * family quirk the request builders already enforce.
 */

import type { AiProtocol } from './registry'

export const THINKING_LEVELS = ['none', 'low', 'medium', 'high', 'xhigh', 'max'] as const

export type ThinkingEffort = (typeof THINKING_LEVELS)[number]

export function isThinkingEffort(value: unknown): value is ThinkingEffort {
  return typeof value === 'string' && (THINKING_LEVELS as readonly string[]).includes(value)
}

/** one family's effort surface: which canonical levels it accepts */
export interface ThinkingShape {
  kind: 'openai-effort' | 'anthropic-adaptive' | 'gemini-level' | 'toggle' | 'qwen-budget'
  levels: readonly ThinkingEffort[]
}

/** GPT-6.x tiers carry the two upper rungs; 5.x tops out at high (models page). */
const OPENAI_SIX_LEVELS: readonly ThinkingEffort[] = [
  'none',
  'low',
  'medium',
  'high',
  'xhigh',
  'max',
]
const OPENAI_FIVE_LEVELS: readonly ThinkingEffort[] = ['none', 'low', 'medium', 'high']
/** o-series models document low/medium/high only — `none` is not a value they take */
const OPENAI_O_SERIES_LEVELS: readonly ThinkingEffort[] = ['low', 'medium', 'high']
/**
 * xAI's grok takes the OpenAI-style effort field and cannot disable reasoning.
 * Per docs.x.ai "Reasoning" (2026-11-03): grok-4.7/4.6/4.5 take
 * low/medium/high, and xhigh is available on grok-4.6 and later (on 4.5 it is
 * clamped to high). The default is high.
 */
const GROK_LEVELS: readonly ThinkingEffort[] = ['low', 'medium', 'high']
const GROK_XHIGH_LEVELS: readonly ThinkingEffort[] = ['low', 'medium', 'high', 'xhigh']
/** the rungs only the Responses API takes: chat-completions gateways 400 on them */
const OPENAI_RESPONSES_ONLY: readonly ThinkingEffort[] = ['xhigh', 'max']
/** the 6.x Sol and Astra tiers refuse `none` outright (registry, live-verified 2026-10-27) */
const NO_NONE = /(^|\/)gpt-6(?:\.\d+)?-(?:astra|sol)$/i
/**
 * Anthropic adaptive thinking arrived with the 4.6 generation: 4.5 and older
 * reject `{ type: 'adaptive' }` with a 400, so those ids carry no knob at all.
 * 4.6 takes low..high; the 5.x lineup extends that past high (models overview).
 * Neither offers `none`: omitting `thinking` still runs adaptive there, and
 * `{ type: 'disabled' }` is a 400 on the newest ones, so the vendor default is
 * the only "off" those ids have.
 */
const ANTHROPIC_FIVE_LEVELS: readonly ThinkingEffort[] = ['low', 'medium', 'high', 'xhigh', 'max']
const ANTHROPIC_FOUR_SIX_LEVELS: readonly ThinkingEffort[] = ['low', 'medium', 'high']
/** Gemini 3.x/4.x `thinkingLevel` bars: the vendor lists low, medium and high */
const GEMINI_LEVELS: readonly ThinkingEffort[] = ['low', 'medium', 'high']

/**
 * The model's family effort surface. `protocol` is the route the request will
 * actually take: the two upper OpenAI rungs exist on the Responses API only, so
 * the same model id behind a chat-completions gateway offers just low..high.
 */
export function thinkingShapeFor(model: string, protocol?: AiProtocol): ThinkingShape {
  const id = model.toLowerCase()
  if (/(^|\/)(gpt-[5-9]|o[134](-|$))/.test(id)) {
    const six = /(^|\/)gpt-6/.test(id)
    const oSeries = /(^|\/)o[134](-|$)/.test(id)
    let levels = six ? OPENAI_SIX_LEVELS : oSeries ? OPENAI_O_SERIES_LEVELS : OPENAI_FIVE_LEVELS
    if (protocol !== 'openai-responses') {
      levels = levels.filter((l) => !OPENAI_RESPONSES_ONLY.includes(l))
    }
    if (NO_NONE.test(model)) levels = levels.filter((l) => l !== 'none')
    return { kind: 'openai-effort', levels }
  }
  if (/(^|\/)claude/.test(id)) {
    // the family spelling is optional: the model field is free text
    const version = /(?:^|\/)(?:claude-)?(?:(?:opus|sonnet|haiku|fable)-)?(\d+)(?:[.-](\d+))?/.exec(
      id,
    )
    const major = version ? Number(version[1]) : 0
    const minor = version && version[2] !== undefined ? Number(version[2]) : 0
    // 5.x extends adaptive effort past high; 4.6 introduced it, older ids 400
    if (major >= 5) return { kind: 'anthropic-adaptive', levels: ANTHROPIC_FIVE_LEVELS }
    if (major === 4 && minor >= 6) {
      return { kind: 'anthropic-adaptive', levels: ANTHROPIC_FOUR_SIX_LEVELS }
    }
    return { kind: 'anthropic-adaptive', levels: [] }
  }
  if (/(^|\/)gemini-[3-9]/.test(id)) return { kind: 'gemini-level', levels: GEMINI_LEVELS }
  // Ant Ling 3.0/3.1: thinking.type enabled|disabled (default enabled) — the
  // API reference documents the switch, no levels
  if (/(^|\/)ling-/.test(id) || /(^|\/)ring-/.test(id))
    return { kind: 'toggle', levels: ['none', 'medium'] }
  // iFlytek Spark X2+ controls its thinking mode through a thinking parameter
  if (/(^|\/)spark-x/.test(id)) return { kind: 'toggle', levels: ['none', 'medium'] }
  // MiMo V2.6: thinking.type enabled/disabled, on by default
  if (/(^|\/)mimo-/.test(id)) return { kind: 'toggle', levels: ['none', 'medium'] }
  // DashScope Qwen 3.x: enable_thinking plus a thinking_budget token cap
  // (per-model ranges; these picks sit inside the band the guide cites)
  if (/(^|\/)qwen\d/.test(id))
    return { kind: 'qwen-budget', levels: ['none', 'low', 'medium', 'high'] }
  // xAI's grok takes the OpenAI-style effort field on the current lineup and
  // states reasoning cannot be disabled; xhigh exists from grok-4.6 on.
  if (/(^|\/)grok-/.test(id)) {
    const version = /(?:^|\/)grok-(\d+)(?:[.-](\d+))?/.exec(id)
    const major = version ? Number(version[1]) : 0
    const minor = version && version[2] !== undefined ? Number(version[2]) : 0
    const levels = major > 4 || (major === 4 && minor >= 6) ? GROK_XHIGH_LEVELS : GROK_LEVELS
    return { kind: 'openai-effort', levels }
  }
  // Doc-only families (GLM, Doubao, MiniMax, DeepSeek, Kimi, Hunyuan, Mistral)
  // deliberately carry no knob until a live probe confirms their surface.
  return { kind: 'openai-effort', levels: [] }
}

/** The stored pick for this model, only when the family (and route) accepts
 * it. A pick a family does not list (an old file, a changed model id, a pick
 * from the Responses API replayed on a chat-completions gateway) resolves to
 * nothing sent rather than a guessed nearest rung. */
export function resolveThinkingEffort(
  config: {
    model: string
    thinking?: Record<string, string> | undefined
  },
  protocol?: AiProtocol,
): ThinkingEffort | undefined {
  const raw = config.thinking?.[config.model]
  if (!isThinkingEffort(raw)) return undefined
  const shape = thinkingShapeFor(config.model, protocol)
  return shape.levels.includes(raw) ? raw : undefined
}

/** `reasoning.effort` (Responses) / `reasoning_effort` (Chat Completions) for
 * the families that speak the OpenAI effort field; undefined = send nothing. */
export function openAiEffortFor(
  config: {
    model: string
    thinking?: Record<string, string> | undefined
  },
  protocol?: AiProtocol,
): string | undefined {
  if (thinkingShapeFor(config.model, protocol).kind !== 'openai-effort') return undefined
  return resolveThinkingEffort(config, protocol)
}

/**
 * Anthropic adaptive-thinking request fields. The record itself carries no
 * level (`{ type: 'adaptive' }`); the pick rides the separate top-level
 * `output_config.effort`, and an `effort` inside `thinking` is rejected as an
 * invalid request. Absent fields = the vendor default. Only 4.6+ ids are
 * adaptive at all; 4.5 and older resolve to no knob, so nothing is sent for
 * them (their own "off" already is the default).
 */
export function anthropicThinkingFields(config: {
  model: string
  thinking?: Record<string, string> | undefined
}): { thinking: { type: 'adaptive' }; effort: string } | undefined {
  if (thinkingShapeFor(config.model).kind !== 'anthropic-adaptive') return undefined
  const level = resolveThinkingEffort(config)
  if (level === undefined || level === 'none') return undefined
  return { thinking: { type: 'adaptive' }, effort: level }
}

/** Gemini 3/4.x `generationConfig.thinkingConfig.thinkingLevel`; absent = dynamic (the default). */
export function geminiThinkingLevelFor(config: {
  model: string
  thinking?: Record<string, string> | undefined
}): string | undefined {
  if (thinkingShapeFor(config.model).kind !== 'gemini-level') return undefined
  return resolveThinkingEffort(config)
}

/** The on/off families (Ling, Spark, MiMo) and Qwen's budget merge extra body
 * fields through the OpenAI-compatible protocol's bodyExtras; undefined = send
 * nothing. */
export function toggleThinkingBodyFor(config: {
  model: string
  thinking?: Record<string, string> | undefined
}): Record<string, unknown> | undefined {
  const kind = thinkingShapeFor(config.model).kind
  const level = resolveThinkingEffort(config)
  if (level === undefined) return undefined
  if (kind === 'toggle') return { thinking: { type: level === 'none' ? 'disabled' : 'enabled' } }
  if (kind === 'qwen-budget') {
    if (level === 'none') return { enable_thinking: false }
    const budget =
      level === 'low' ? 4096 : level === 'medium' ? 8192 : level === 'high' ? 16384 : undefined
    if (budget === undefined) return undefined
    return { enable_thinking: true, thinking_budget: budget }
  }
  return undefined
}
