import { describe, expect, it } from 'vitest'
import {
  deriveDeckProgressView,
  type DeckProgressSnapshot,
} from '../src/renderer/ai/deck-progress-view'
import type { TFunc } from '../src/renderer/i18n/locale'

const t: TFunc = (key, params) => (params ? `${key}:${JSON.stringify(params)}` : key)

function pages(items: Array<'done' | 'error' | 'running'>): DeckProgressSnapshot['pages'] {
  return {
    label: 'Generating slides',
    done: items.filter((s) => s === 'done').length,
    total: items.length,
    status: 'running',
    summary: 'Generating page 2/3',
    items: items.map((status, i) => ({ title: `P${i + 1}`, status })),
  }
}

describe('deriveDeckProgressView', () => {
  it('reads a zero-page done event as a failure, not as "done — 0 slides"', () => {
    const view = deriveDeckProgressView(
      {
        pages: pages(['error', 'error', 'error']),
        isDone: true,
        finalTotal: 0,
        doneSummary: 'All 3 pages failed',
        doneOutcome: 'failed',
      },
      t,
    )
    expect(view.head).toEqual({ text: 'aiProgressFailed', tone: 'error' })
    const step = view.steps.find((s) => s.key === 'pages')!
    expect(step.stepStatus).toBe('error')
    expect(step.label).toBe('All 3 pages failed')
    expect(step.label).not.toContain('Generating')
  })

  it('marks the run failed even when no page item carries an error flag', () => {
    const view = deriveDeckProgressView(
      { pages: pages(['running']), isDone: true, finalTotal: 0, doneOutcome: 'failed' },
      t,
    )
    expect(view.head.tone).toBe('error')
    expect(view.steps[0]!.stepStatus).toBe('error')
  })

  it('keeps the success line for a completed run', () => {
    const view = deriveDeckProgressView(
      { pages: pages(['done', 'done']), isDone: true, finalTotal: 2, doneSummary: 'Done 2' },
      t,
    )
    expect(view.head).toEqual({ text: 'aiProgressDone:{"n":2}', tone: 'done' })
    expect(view.steps[0]).toMatchObject({
      stepStatus: 'done',
      label: 'Generating slidesaiPagesSuffix:{"n":2}',
    })
  })

  it('shows the stop summary, not the success line, for a cancelled run', () => {
    const view = deriveDeckProgressView(
      {
        pages: pages(['done', 'running']),
        isDone: true,
        finalTotal: 1,
        doneSummary: 'Stopped (kept 1)',
        doneOutcome: 'cancelled',
      },
      t,
    )
    expect(view.head).toEqual({ text: 'Stopped (kept 1)', tone: 'done' })
  })

  it('partial success stays a done head with an error-flagged pages step', () => {
    const view = deriveDeckProgressView(
      { pages: pages(['done', 'error']), isDone: true, finalTotal: 1, doneSummary: 'Done 1' },
      t,
    )
    expect(view.head.tone).toBe('done')
    expect(view.steps[0]!.stepStatus).toBe('error')
  })

  it('spins with the live summary while pages are still landing', () => {
    const view = deriveDeckProgressView({ pages: pages(['done', 'running', 'running']) }, t)
    expect(view.head).toEqual({ text: 'aiProgressTitle', tone: 'running' })
    expect(view.steps[0]).toMatchObject({ stepStatus: 'running', label: 'Generating page 2/3' })
  })
})
