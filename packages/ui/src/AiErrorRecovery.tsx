import React from 'react'

import { IconClose, IconEdit, IconRetry, IconWarning } from './icons'
import { AIError, classifyError, type AIErrorCode } from '@genoffice/chat-runtime/errors'

export interface AiErrorRecoveryProps {
  error: AIError | Error | string | null | undefined
  /** Optional title override; falls back to the default for each error code. */
  title?: string
  /** Retry click handler — only shown when the error is `retryable`. */
  onRetry?: () => void
  /** Edit click handler (e.g. re-open the composer with the failed prompt). */
  onEdit?: () => void
  /** Dismiss / cancel click handler — always shown when provided. */
  onDismiss?: () => void
  retryLabel?: string
  editLabel?: string
  dismissLabel?: string
  className?: string
}

const TITLES: Record<AIErrorCode, string> = {
  NOT_CONFIGURED: 'AI not configured',
  NO_MODEL: 'No model selected',
  NETWORK: 'Network error',
  TIMEOUT: 'Request timed out',
  CANCELLED: 'Run cancelled',
  TOOL_FAILED: 'Tool failed',
  WEB_UNSUPPORTED: 'Not supported in web',
  PROVIDER: 'Provider error',
  INTERNAL: 'Something went wrong',
}

const RETRY_LABELS: Record<AIErrorCode, string> = {
  NOT_CONFIGURED: 'Open settings',
  NO_MODEL: 'Choose model',
  NETWORK: 'Retry',
  TIMEOUT: 'Try again',
  CANCELLED: 'Restart',
  TOOL_FAILED: 'Retry',
  WEB_UNSUPPORTED: 'Open desktop',
  PROVIDER: 'Retry',
  INTERNAL: 'Retry',
}

/**
 * Inline error card that surfaces the normalised `AIErrorCode` from
 * `@genoffice/chat-runtime/errors`. Each app renders the same component;
 * the buttons call back into the host (retry the run, edit the
 * composer, or dismiss) so the host owns the actual action.
 */
export function AiErrorRecovery({
  error,
  title,
  onRetry,
  onEdit,
  onDismiss,
  retryLabel,
  editLabel = 'Edit prompt',
  dismissLabel = 'Dismiss',
  className,
}: AiErrorRecoveryProps): React.JSX.Element | null {
  if (!error) return null
  const normalised = error instanceof AIError ? error : classifyError(error)
  const heading = title ?? TITLES[normalised.code] ?? 'Error'
  const retryText = retryLabel ?? RETRY_LABELS[normalised.code] ?? 'Retry'
  const retryPrimary = normalised.retryable
  return (
    <div
      className={['ai-error-recovery', className].filter(Boolean).join(' ')}
      role="alert"
      aria-label={heading}
    >
      <div className="ai-error-recovery-title">
        <IconWarning size={14} />
        <span>{heading}</span>
      </div>
      <div className="ai-error-recovery-message">{normalised.message}</div>
      {(onRetry || onEdit || onDismiss) && (
        <div className="ai-error-recovery-actions">
          {onRetry && (
            <button
              type="button"
              className="ai-error-recovery-action"
              data-primary={retryPrimary ? 'true' : 'false'}
              onClick={() => onRetry()}
            >
              <IconRetry size={12} /> {retryText}
            </button>
          )}
          {onEdit && (
            <button
              type="button"
              className="ai-error-recovery-action"
              onClick={() => onEdit()}
            >
              <IconEdit size={12} /> {editLabel}
            </button>
          )}
          {onDismiss && (
            <button
              type="button"
              className="ai-error-recovery-action"
              onClick={() => onDismiss()}
            >
              <IconClose size={12} /> {dismissLabel}
            </button>
          )}
        </div>
      )}
    </div>
  )
}
