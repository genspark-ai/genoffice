import React from 'react'

import { IconCheck, IconClose, IconWarning } from './icons'
import { summarizeChangePlan } from '@genoffice/chat-runtime/change-plan'
import type { ChatChangePlan } from '@genoffice/chat-runtime/types'

export interface AiChangeSummaryProps {
  plan: ChatChangePlan
  /**
   * Optional override for the bullet renderer. When omitted, falls back to
   * `summarizeChangePlan()` from `@genoffice/chat-runtime/change-plan`,
   * which produces one bullet per op.
   */
  previewRenderer?: (op: ChatChangePlan['ops'][number]) => string
  /**
   * Optional rich preview slot. When provided, renders below the bullet
   * list (e.g. an XLSX preview table or a doc diff). When omitted, only
   * the bullets show.
   */
  preview?: React.ReactNode
  /** Apply button label. Default: 'Apply'. */
  applyLabel?: string
  /** Reject button label. Default: 'Reject'. */
  rejectLabel?: string
  /** Called when the user clicks Apply. */
  onApply?: () => void
  /** Called when the user clicks Reject. */
  onReject?: () => void
  /** Optional className passthrough. */
  className?: string
}

/**
 * Card showing the most recent `ChatChangePlan`: title, warnings, bullet
 * preview, and (when handlers are wired) apply / reject actions. Each app
 * can override the bullet renderer or provide a rich `preview` slot for
 * app-specific visualisation (e.g. workbook preview table).
 */
export function AiChangeSummary({
  plan,
  previewRenderer,
  preview,
  applyLabel = 'Apply',
  rejectLabel = 'Reject',
  onApply,
  onReject,
  className,
}: AiChangeSummaryProps): React.JSX.Element {
  const bullets = summarizeChangePlan(plan, previewRenderer)
  return (
    <div
      className={['ai-change-summary', className].filter(Boolean).join(' ')}
      data-app={plan.app}
      role="region"
      aria-label={`Change plan: ${plan.title}`}
    >
      <div className="ai-change-summary-title">
        <IconCheck size={14} />
        <span>{plan.title}</span>
      </div>
      {bullets.length > 0 && (
        <ul className="ai-change-summary-bullets">
          {bullets.map((b, i) => (
            <li key={i}>{b}</li>
          ))}
        </ul>
      )}
      {preview}
      {plan.warnings && plan.warnings.length > 0 && (
        <div className="ai-change-summary-warnings">
          {plan.warnings.map((w, i) => (
            <div key={i}>
              <IconWarning size={12} /> {w}
            </div>
          ))}
        </div>
      )}
      {(onApply || onReject) && (
        <div className="ai-change-summary-actions">
          {onApply && (
            <button
              type="button"
              className="ai-change-summary-action"
              data-variant="primary"
              onClick={() => onApply()}
            >
              <IconCheck size={12} /> {applyLabel}
            </button>
          )}
          {onReject && (
            <button
              type="button"
              className="ai-change-summary-action"
              data-variant="ghost"
              onClick={() => onReject()}
            >
              <IconClose size={12} /> {rejectLabel}
            </button>
          )}
        </div>
      )}
    </div>
  )
}
