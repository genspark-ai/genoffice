import React from 'react'
import type { Lang } from '@genoffice/i18n'
import type { AgentTokenUsage } from '@genoffice/agent-core'
import { AI_USAGE_LABELS } from './strings-ai-usage'
import './ai-usage.css'

/** usage figures a chip can render (structural: agent-core's AgentTokenUsage satisfies it) */
export type AiTokenUsageView = AgentTokenUsage

/** 1234 → "1.2k", 1234567 → "1.2M"; smaller counts stay verbatim. */
export function formatTokenCount(n: number): string {
  if (n >= 1_000_000) return `${(n / 1_000_000).toFixed(1).replace(/\.0$/, '')}M`
  if (n >= 1_000) return `${(n / 1_000).toFixed(1).replace(/\.0$/, '')}k`
  return String(n)
}

/** Add one turn's reported usage into an accumulator; sides the provider never
 * reported stay absent instead of turning into fake zeros. */
export function addTokenUsage(
  prev: AgentTokenUsage | undefined,
  delta: AgentTokenUsage,
): AgentTokenUsage {
  const out: AgentTokenUsage = { ...(prev ?? {}) }
  if (delta.inputTokens !== undefined) out.inputTokens = (out.inputTokens ?? 0) + delta.inputTokens
  if (delta.outputTokens !== undefined)
    out.outputTokens = (out.outputTokens ?? 0) + delta.outputTokens
  return out
}

function usageText(usage: AgentTokenUsage): string {
  const parts: string[] = []
  if (usage.inputTokens !== undefined) parts.push(`↑${formatTokenCount(usage.inputTokens)}`)
  if (usage.outputTokens !== undefined) parts.push(`↓${formatTokenCount(usage.outputTokens)}`)
  return parts.join(' ')
}

/**
 * Small numeric chip showing tokens billed for a reply (or a session total).
 * Renders nothing when the provider reported nothing — usage is optional data,
 * never a fake zero.
 */
export function AiUsageChip({
  usage,
  session,
  lang,
}: {
  readonly usage: AgentTokenUsage
  /** session-total styling (sum icon instead of the per-reply arrows) */
  readonly session?: boolean | undefined
  readonly lang: Lang
}): React.JSX.Element | null {
  const text = usageText(usage)
  if (!text) return null
  const title = session ? AI_USAGE_LABELS[lang].sessionTitle : AI_USAGE_LABELS[lang].runTitle
  return (
    <span className="ai-usage-chip" title={title}>
      {session ? 'Σ ' : ''}
      {text}
    </span>
  )
}
