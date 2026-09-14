import React from 'react'

import { IconCheck, IconStop, IconTool, IconWarning } from './icons'
import type { ChatToolCallRecord } from '@genoffice/chat-runtime/types'

export interface AiToolTimelineProps {
  tools: ChatToolCallRecord[]
  /** Optional renderer used for tool-specific labels (e.g. "Insert column at C"). */
  formatName?: (tool: ChatToolCallRecord) => string
  /** Optional summary line shown under each tool row (e.g. tool display text). */
  summaryOf?: (tool: ChatToolCallRecord) => string | undefined
  className?: string
}

/**
 * Vertical list of tool calls observed during the current run. Each row
 * shows the tool name, current status, and (when known) a duration so
 * the user can see what's happening in long multi-tool runs.
 */
export function AiToolTimeline({
  tools,
  formatName,
  summaryOf,
  className,
}: AiToolTimelineProps): React.JSX.Element | null {
  if (!tools.length) return null
  return (
    <div
      className={['ai-tool-timeline', className].filter(Boolean).join(' ')}
      aria-label={`${tools.length} tool call${tools.length === 1 ? '' : 's'}`}
    >
      {tools.map(tool => {
        const name = formatName ? formatName(tool) : tool.name
        const summary = summaryOf?.(tool)
        const duration = formatDuration(tool.startedAt, tool.finishedAt)
        return (
          <div key={tool.id} className="ai-tool-row" data-status={tool.status}>
            <span className="ai-tool-row-icon">{iconFor(tool.status)}</span>
            <span className="ai-tool-row-name">{name}</span>
            <span className="ai-tool-row-status">
              {tool.status === 'running' ? 'running…' : duration || (tool.status === 'error' ? 'failed' : 'ok')}
            </span>
            {summary && <span className="ai-tool-row-summary">{summary}</span>}
          </div>
        )
      })}
    </div>
  )
}

function iconFor(status: ChatToolCallRecord['status']): React.ReactNode {
  if (status === 'running') return <IconTool size={14} />
  if (status === 'error') return <IconWarning size={14} />
  if (status === 'executed') return <IconCheck size={14} />
  return <IconStop size={14} />
}

function formatDuration(start: number, finish: number | undefined): string {
  if (!finish) return ''
  const ms = finish - start
  if (ms < 1000) return `${ms} ms`
  if (ms < 60_000) return `${(ms / 1000).toFixed(1)} s`
  return `${Math.floor(ms / 60_000)}m ${Math.round((ms % 60_000) / 1000)}s`
}
