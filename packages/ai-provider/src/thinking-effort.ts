/**
 * Per-model thinking-effort resolution: what the user picked in settings,
 * translated into the one request field each model family actually accepts.
 *
 * The levels follow OpenAI's current ladder (the widest one): `none` turns
 * reasoning off, `xhigh`/`max` are the upper tiers most 2026 lineups share.
 * Every other family maps onto that ladder with its own shape, and a family
 * without any effort parameter (MiniMax — live-verified rejects — Mistral,
 * Ling, Spark, LongCat, MiMo, Hunyuan, codex CLI) resolves to nothing sent,
 * i.e. today's behaviour.
 *
 * Resolution order mirrors the temperature rule: a user's stored pick wins,
 * an absent pick means the vendor default, and nothing here overrides a
 * family quirk the request builders already enforce.
 */

export const THINKING_LEVELS = ['none', 'low', 'medium', 'high', 'xhigh', 'max'] as const

export type ThinkingEffort = (typeof THINKING_LEVELS)[number]

export function isThinkingEffort(value: unknown): value is ThinkingEffort {
  return typeof value === 'string' && (THINKING_LEVELS as readonly string[]).includes(value)
}

/** one family's effort surface: which canonical levels it accepts */
export interface ThinkingShape {
  kind: 'openai-effort' | 'anthropic-adaptive' | 'gemini-level' | 'toggle' | 'qwen-toggle'
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
/** the 6.1 Sol and Astra tiers refuse `none` outright (registry, live-verified 2026-10-27) */
const NO_NONE = /(^|\/)gpt-6-(?:astra|sol)$/i
/** Anthropic's 5.x lineup extends adaptive effort past high (models overview) */
const ANTHROPIC_FIVE_LEVELS: readonly ThinkingEffort[] = [
  'none',
  'low',
  'medium',
  'high',
  'xhigh',
  'max',
]
const ANTHROPIC_FOUR_LEVELS: readonly ThinkingEffort[] = ['none', 'low', 'medium', 'high']

/** the model's family effort surface; empty levels = nothing to send */
export function thinkingShapeFor(model: string): ThinkingShape {
  const id = model.toLowerCase()
  if (/(^|\/)(gpt-[5-9]|o[134](-|$))/.test(id)) {
    const six = /(^|\/)gpt-6/.test(id)
    const base = six ? OPENAI_SIX_LEVELS : OPENAI_FIVE_LEVELS
    return {
      kind: 'openai-effort',
      levels: NO_NONE.test(model) ? base.filter((l) => l !== 'none') : base,
    }
  }
  if (/(^|\/)claude/.test(id)) {
    const five = /(opus-5|sonnet-5|fable-5|haiku-5)/.test(id)
    return {
      kind: 'anthropic-adaptive',
      levels: five ? ANTHROPIC_FIVE_LEVELS : ANTHROPIC_FOUR_LEVELS,
    }
  }
  if (/(^|\/)gemini-[3-9]/.test(id)) return { kind: 'gemini-level', levels: ['low', 'high'] }
  // DeepSeek V4 and Kimi K3 take the same top-level reasoning_effort field with
  // their own ladders (DeepSeek's encoding.py: low/high/max presets or an
  // integer; Kimi's K3 docs: low/high/max, default max). V4 Flash ships
  // non-thinking by default, so omitting the field is already "off" — no none.
  if (/(^|\/)deep-?seek-/.test(id)) return { kind: 'openai-effort', levels: ['low', 'high', 'max'] }
  if (/(^|\/)kimi-k3/.test(id)) return { kind: 'openai-effort', levels: ['low', 'high', 'max'] }
  // GLM-5.2+ configures reasoning_effort across the full ladder (its docs alias
  // low/medium→high and xhigh→max, and none/minimal disable thinking); 5.3 is
  // always-on, so it offers everything but none.
  if (/(^|\/)glm-5\.3/.test(id))
    return { kind: 'openai-effort', levels: ['low', 'medium', 'high', 'xhigh', 'max'] }
  if (/(^|\/)glm-/.test(id))
    return { kind: 'openai-effort', levels: ['none', 'low', 'medium', 'high', 'xhigh', 'max'] }
  // Doubao Seed 2.1: the Ark gateway accepts the full effort ladder and maps
  // internally (none/minimal stop thinking, xhigh/max fold to high on 2.1);
  // the hard on/off stays available through thinking{type} which we do not
  // need once effort carries the choice.
  if (/(^|\/)doubao-seed-2-1/.test(id)) return { kind: 'openai-effort', levels: OPENAI_SIX_LEVELS }
  // Hunyuan Hy: reasoning_effort with low/high documented (default high) and
  // a thinking{type} switch on top; hy4-preview also thinks by default.
  if (/(^|\/)hy[34](-|$)/.test(id)) return { kind: 'openai-effort', levels: ['low', 'high'] }
  // iFlytek Spark X2+ controls its thinking mode through a thinking parameter
  if (/(^|\/)spark-x/.test(id)) return { kind: 'toggle', levels: ['none', 'medium'] }
  // MiniMax M3/M3.1: reasoning_effort low..max on the compatible route —
  // thinking is always on (a disabled/none attempt answers 400), so no none.
  if (/(^|\/)minimax-m/.test(id))
    return { kind: 'openai-effort', levels: ['low', 'medium', 'high', 'xhigh', 'max'] }
  // MiMo V2.6: thinking.type enabled/disabled, on by default
  if (/(^|\/)mimo-/.test(id)) return { kind: 'toggle', levels: ['none', 'medium'] }
  // DashScope Qwen 3.x commercial tiers: enable_thinking
  if (/(^|\/)qwen\d/.test(id)) return { kind: 'qwen-toggle', levels: ['none', 'medium'] }
  // xAI's grok takes the OpenAI-style effort field on the current lineup
  if (/(^|\/)grok-/.test(id)) return { kind: 'openai-effort', levels: OPENAI_FIVE_LEVELS }
  // Mistral documents reasoning_effort for medium 3.5 (three rungs); the other
  // mistral ids have no documented knob
  if (/(^|\/)mistral-medium/.test(id))
    return { kind: 'openai-effort', levels: ['low', 'medium', 'high'] }
  return { kind: 'openai-effort', levels: [] }
}

/** The stored pick for this model, only when the family accepts it. A pick a
 * family does not list (an old file, a changed model id) resolves to nothing
 * sent rather than a guessed nearest rung. */
export function resolveThinkingEffort(config: {
  model: string
  thinking?: Record<string, string> | undefined
}): ThinkingEffort | undefined {
  const raw = config.thinking?.[config.model]
  if (!isThinkingEffort(raw)) return undefined
  const shape = thinkingShapeFor(config.model)
  return shape.levels.includes(raw) ? raw : undefined
}

/** `reasoning.effort` (Responses) / `reasoning_effort` (Chat Completions) for
 * the families that speak the OpenAI effort field; undefined = send nothing. */
export function openAiEffortFor(config: {
  model: string
  thinking?: Record<string, string> | undefined
}): string | undefined {
  if (thinkingShapeFor(config.model).kind !== 'openai-effort') return undefined
  return resolveThinkingEffort(config)
}

/** Anthropic's adaptive-thinking record; `none` resolves to undefined, which
 * is the vendor's own "no thinking block" default. */
export function anthropicThinkingFor(config: {
  model: string
  thinking?: Record<string, string> | undefined
}): { type: 'adaptive'; effort: string } | undefined {
  if (thinkingShapeFor(config.model).kind !== 'anthropic-adaptive') return undefined
  const level = resolveThinkingEffort(config)
  if (level === undefined || level === 'none') return undefined
  return { type: 'adaptive', effort: level }
}

/** Gemini 3/4.x `generationConfig.thinkingLevel`; absent = dynamic (the default). */
export function geminiThinkingLevelFor(config: {
  model: string
  thinking?: Record<string, string> | undefined
}): string | undefined {
  if (thinkingShapeFor(config.model).kind !== 'gemini-level') return undefined
  return resolveThinkingEffort(config)
}

/** The on/off families (GLM, Doubao, Qwen) merge extra body fields through the
 * OpenAI-compatible protocol's bodyExtras; undefined = send nothing. */
export function toggleThinkingBodyFor(config: {
  model: string
  thinking?: Record<string, string> | undefined
}): Record<string, unknown> | undefined {
  const kind = thinkingShapeFor(config.model).kind
  const level = resolveThinkingEffort(config)
  if (level === undefined) return undefined
  if (kind === 'toggle') return { thinking: { type: level === 'none' ? 'disabled' : 'enabled' } }
  if (kind === 'qwen-toggle') return { enable_thinking: level !== 'none' }
  return undefined
}
