/**
 * Render smoke tests for the M3 shared AI runtime primitives.
 *
 * Uses `react-dom/server.renderToStaticMarkup` so we don't need jsdom or
 * a full testing-library setup; we just verify the components emit the
 * expected class names, status pills, and recovery actions.
 */
import { describe, expect, it } from 'vitest'
import { renderToStaticMarkup } from 'react-dom/server'
import { createElement } from 'react'

import { AiRunHeader } from '../src/AiRunHeader'
import { AiProviderBadge } from '../src/AiProviderBadge'
import { AiAttachmentStrip } from '../src/AiAttachmentStrip'
import { AiToolTimeline } from '../src/AiToolTimeline'
import { AiChangeSummary } from '../src/AiChangeSummary'
import { AiErrorRecovery } from '../src/AiErrorRecovery'
import { AIError } from '@genoffice/chat-runtime/errors'
import { normalizeChangePlan } from '@genoffice/chat-runtime/change-plan'
import type { ChatToolCallRecord } from '@genoffice/chat-runtime/types'

describe('AiRunHeader', () => {
  it('renders nothing for the idle state', () => {
    expect(renderToStaticMarkup(createElement(AiRunHeader, { status: 'idle' }))).toBe('')
  })

  it('shows the status pill + model badge + run id when running', () => {
    const html = renderToStaticMarkup(
      createElement(AiRunHeader, {
        status: 'streaming',
        runId: 'run-abc',
        model: 'MiniMax M3',
        onStop: () => undefined,
      }),
    )
    expect(html).toContain('ai-run-header-pill')
    expect(html).toContain('data-status="streaming"')
    expect(html).toContain('ai-provider-badge')
    expect(html).toContain('MiniMax M3')
    expect(html).toContain('#run-abc')
    expect(html).toContain('ai-run-header-stop')
  })

  it('hides the stop button when no onStop is provided', () => {
    const html = renderToStaticMarkup(
      createElement(AiRunHeader, { status: 'running' }),
    )
    expect(html).not.toContain('ai-run-header-stop')
  })

  it('hides the stop button when status is terminal (done/error)', () => {
    const html = renderToStaticMarkup(
      createElement(AiRunHeader, { status: 'done', onStop: () => undefined }),
    )
    expect(html).not.toContain('ai-run-header-stop')
  })
})

describe('AiProviderBadge', () => {
  it('renders the label or a default', () => {
    expect(renderToStaticMarkup(createElement(AiProviderBadge, { label: 'MiniMax M3' }))).toContain('MiniMax M3')
    expect(renderToStaticMarkup(createElement(AiProviderBadge))).toContain('AI')
  })
})

describe('AiAttachmentStrip', () => {
  it('renders nothing when there are no attachments', () => {
    expect(renderToStaticMarkup(createElement(AiAttachmentStrip, { attachments: [] }))).toBe('')
  })

  it('renders one chip per attachment with size formatting', () => {
    const html = renderToStaticMarkup(
      createElement(AiAttachmentStrip, {
        attachments: [
          { id: 'a1', name: 'spec.md', type: 'text/markdown', size: 512 },
          { id: 'a2', name: 'chart.png', type: 'image/png', size: 1_500_000 },
        ],
      }),
    )
    expect(html).toContain('ai-attachment-chip')
    expect(html).toContain('spec.md')
    expect(html).toContain('chart.png')
    expect(html).toContain('512 B')
    expect(html).toContain('1.4 MB')
  })
})

describe('AiToolTimeline', () => {
  const now = Date.now()
  const tools: ChatToolCallRecord[] = [
    { id: 't1', name: 'search_text', input: {}, status: 'executed', startedAt: now - 1500, finishedAt: now },
    { id: 't2', name: 'edit_block', input: {}, status: 'running', startedAt: now },
    { id: 't3', name: 'add_image', input: {}, status: 'error', isError: true, startedAt: now - 800, finishedAt: now - 100 },
  ]

  it('renders one row per tool with the right status', () => {
    const html = renderToStaticMarkup(createElement(AiToolTimeline, { tools }))
    expect(html).toContain('ai-tool-row')
    expect(html.match(/data-status=/g)?.length).toBe(3)
    expect(html).toContain('search_text')
    expect(html).toContain('edit_block')
    expect(html).toContain('add_image')
  })

  it('returns null when there are no tools', () => {
    expect(renderToStaticMarkup(createElement(AiToolTimeline, { tools: [] }))).toBe('')
  })

  it('honours the formatName and summaryOf callbacks', () => {
    const html = renderToStaticMarkup(
      createElement(AiToolTimeline, {
        tools,
        formatName: t => `tool:${t.name}`,
        summaryOf: () => 'summary line',
      }),
    )
    expect(html).toContain('tool:search_text')
    expect(html).toContain('summary line')
  })
})

describe('AiChangeSummary', () => {
  const plan = normalizeChangePlan({
    app: 'sheets',
    title: 'Insert column',
    summary: 'Insert a SUM column at C',
    ops: [{ op: 'insert-column', at: 'C' }, { op: 'set-cell', at: 'C1', value: 'Total' }],
  })

  it('renders the title and one bullet per op', () => {
    const html = renderToStaticMarkup(createElement(AiChangeSummary, { plan }))
    expect(html).toContain('Insert column')
    expect(html).toContain('ai-change-summary-bullets')
    expect(html.match(/<li/g)?.length).toBe(2)
  })

  it('renders Apply/Reject buttons when handlers are provided', () => {
    const html = renderToStaticMarkup(
      createElement(AiChangeSummary, {
        plan,
        onApply: () => undefined,
        onReject: () => undefined,
      }),
    )
    expect(html).toContain('Apply')
    expect(html).toContain('Reject')
  })

  it('omits the action bar when no handlers are wired', () => {
    const html = renderToStaticMarkup(createElement(AiChangeSummary, { plan }))
    expect(html).not.toContain('ai-change-summary-actions')
  })
})

describe('AiErrorRecovery', () => {
  it('returns null when there is no error', () => {
    expect(renderToStaticMarkup(createElement(AiErrorRecovery, { error: null }))).toBe('')
  })

  it('shows the right title + retry label for a WEB_UNSUPPORTED error', () => {
    const err = new AIError('WEB_UNSUPPORTED', 'Channel not supported on web', { channel: 'ai:doc-write' })
    const html = renderToStaticMarkup(
      createElement(AiErrorRecovery, {
        error: err,
        onRetry: () => undefined,
      }),
    )
    expect(html).toContain('Not supported in web')
    expect(html).toContain('Open desktop')
    expect(html).toContain('Channel not supported on web')
  })

  it('classifies a free-string error and shows a generic retry', () => {
    const html = renderToStaticMarkup(
      createElement(AiErrorRecovery, {
        error: 'something exploded',
        onRetry: () => undefined,
      }),
    )
    expect(html).toContain('Something went wrong')
    expect(html).toContain('Retry')
    expect(html).toContain('something exploded')
  })

  it('marks the retry button as primary only when the error is retryable', () => {
    const retryable = renderToStaticMarkup(
      createElement(AiErrorRecovery, {
        error: new AIError('TIMEOUT', 'timed out'),
        onRetry: () => undefined,
      }),
    )
    expect(retryable).toContain('data-primary="true"')

    const nonRetryable = renderToStaticMarkup(
      createElement(AiErrorRecovery, {
        error: new AIError('WEB_UNSUPPORTED', 'nope'),
        onRetry: () => undefined,
      }),
    )
    expect(nonRetryable).toContain('data-primary="false"')
  })
})
