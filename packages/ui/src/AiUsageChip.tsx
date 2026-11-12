import React from 'react'
import type { Lang } from '@genoffice/i18n'
import { addTokenUsage, type AgentTokenUsage } from '@genoffice/agent-core'
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

export { addTokenUsage }

/** whole percent of the input served from the prompt cache; null when the provider reports no cache figures */
export function cacheHitPercent(usage: AgentTokenUsage): number | null {
  if (usage.cacheReadTokens === undefined || !usage.inputTokens) return null
  return Math.round((usage.cacheReadTokens / usage.inputTokens) * 100)
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
  const labels = AI_USAGE_LABELS[lang]
  const hit = cacheHitPercent(usage)
  const title =
    (session ? labels.sessionTitle : labels.runTitle) +
    (hit === null ? '' : ` · ${labels.cacheHit.replace('{pct}', String(hit))}`)
  return (
    <span className="ai-usage-chip" title={title}>
      {session ? 'Σ ' : ''}
      {text}
    </span>
  )
}
