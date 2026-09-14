import React from 'react'

export interface AiProviderBadgeProps {
  /** Human-readable label ("MiniMax M3"). Falls back to "AI" when omitted. */
  label?: string
  /** Optional aria-label override. */
  ariaLabel?: string
  className?: string
}

/**
 * Small chip identifying the active provider/model. Used inline next to
 * the run header and inside chat bubbles; intentionally lightweight so
 * each app's chrome can override the look.
 */
export function AiProviderBadge({
  label,
  ariaLabel,
  className,
}: AiProviderBadgeProps): React.JSX.Element {
  return (
    <span
      className={['ai-provider-badge', className].filter(Boolean).join(' ')}
      aria-label={ariaLabel ?? label ?? 'AI'}
      title={label}
    >
      {label ?? 'AI'}
    </span>
  )
}
