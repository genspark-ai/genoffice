import React from 'react'

import { IconStopFilled } from './icons'
import type { ChatRunStatus } from '@genoffice/chat-runtime/types'

export interface AiRunHeaderProps {
  /** Run id displayed as `run-…` after the title. */
  runId?: string
  status: ChatRunStatus
  /** Optional capability label (e.g. "MiniMax M3") — hidden when omitted. */
  model?: string
  /** Optional status pill label override; falls back to the default for each status. */
  statusLabel?: string
  /** Stop button label; only rendered when `onStop` is provided. */
  stopLabel?: string
  /** Stop click handler. When omitted, the button is hidden. */
  onStop?: () => void
  /** Optional className passthrough. */
  className?: string
}

const DEFAULT_LABELS: Record<ChatRunStatus, string> = {
  idle: 'Idle',
  queued: 'Queued',
  running: 'Running',
  streaming: 'Streaming',
  'awaiting-confirm': 'Awaiting confirmation',
  done: 'Done',
  cancelled: 'Cancelled',
  error: 'Error',
}

/**
 * Compact row pinned to the top of the AI chat timeline: status pill,
 * model label, run id, and (when a run is active) a stop button. Replaces
 * the per-app inline copies in Docs / Slides / PDF / XLSX.
 */
export function AiRunHeader(props: AiRunHeaderProps): React.JSX.Element | null {
  const { runId, status, model, statusLabel, stopLabel, onStop, className } = props
  // Hide the row entirely for the terminal `idle` state — there's nothing to show.
  if (status === 'idle') return null
  const label = statusLabel ?? DEFAULT_LABELS[status] ?? status
  const showStop = Boolean(onStop) && (status === 'running' || status === 'streaming' || status === 'queued')
  return (
    <div className={['ai-run-header', className].filter(Boolean).join(' ')} role="status" aria-label={`Run status: ${label}`}>
      <span className="ai-run-header-pill" data-status={status}>{label}</span>
      {model && <span className="ai-provider-badge">{model}</span>}
      {runId && <span className="ai-run-header-id">#{runId}</span>}
      {showStop && (
        <button
          type="button"
          className="ai-run-header-stop"
          aria-label={stopLabel ?? 'Stop run'}
          title={stopLabel ?? 'Stop run'}
          onClick={() => onStop?.()}
        >
          <IconStopFilled size={14} />
        </button>
      )}
    </div>
  )
}
