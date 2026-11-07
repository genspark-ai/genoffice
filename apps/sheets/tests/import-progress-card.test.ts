// @vitest-environment jsdom
/**
 * The delimited import (a .csv/.tsv opened as a workbook) shows how far it is
 * and offers a way out: a 250 MB file takes a while, and the user should see
 * the bar move and be able to walk away.
 *
 * The card is rendered here rather than asserted against the source: a bar
 * that is present but stuck at zero, or a cancel button that never reaches
 * the main process, both look right in the markup and fail in the hand.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it } from 'vitest'

import { ImportProgressCard } from '../src/renderer/ImportProgressCard'
import { importPercent } from '../src/renderer/import-progress'

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
})

interface Rendered {
  readonly percent: string | null
  readonly ariaNow: string | null
  readonly cancelText: string | null
  readonly barWidth: string
  readonly clicks: () => number
}

async function renderCard(
  progress: { phase: 'read' | 'write'; fraction: number } | null,
): Promise<Rendered> {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  let clicks = 0
  await act(async () => {
    root?.render(
      createElement(ImportProgressCard, {
        progress,
        label: 'Opening workbook',
        cancelLabel: 'Cancel',
        onCancel: () => {
          clicks += 1
        },
      }),
    )
  })
  const bar = container.querySelector('.workbook-opening-bar')
  const fill = container.querySelector<HTMLElement>('.workbook-opening-bar-fill')
  return {
    percent: container.querySelector('.workbook-opening-percent')?.textContent ?? null,
    ariaNow: bar?.getAttribute('aria-valuenow') ?? null,
    cancelText: container.querySelector('.workbook-opening-cancel')?.textContent ?? null,
    barWidth: fill?.style.width ?? '',
    clicks: () => clicks,
  }
}

describe('importPercent', () => {
  it('gives the read phase a small slice and the write phase the rest', () => {
    expect(importPercent({ phase: 'read', fraction: 0 })).toBe(0)
    expect(importPercent({ phase: 'read', fraction: 1 })).toBeCloseTo(0.05)
    expect(importPercent({ phase: 'write', fraction: 0 })).toBeCloseTo(0.05)
    expect(importPercent({ phase: 'write', fraction: 1 })).toBe(1)
  })

  it('never leaves the 0..1 range', () => {
    expect(importPercent({ phase: 'read', fraction: -3 })).toBe(0)
    expect(importPercent({ phase: 'write', fraction: 7 })).toBe(1)
  })

  it('only grows', () => {
    const read = importPercent({ phase: 'read', fraction: 0.5 })
    const write = importPercent({ phase: 'write', fraction: 0 })
    expect(write).toBeGreaterThanOrEqual(read)
  })
})

describe('the import progress card', () => {
  it('shows no bar and no cancel button when nothing is importing', async () => {
    const rendered = await renderCard(null)
    expect(rendered.percent).toBe(null)
    expect(rendered.cancelText).toBe(null)
    expect(rendered.ariaNow).toBe(null)
  })

  it('shows the read phase as a small slice of the bar', async () => {
    const rendered = await renderCard({ phase: 'read', fraction: 1 })
    // the whole read phase is worth a few percent, not the whole bar: the
    // read is a pass over the text, the write is the archive, and the
    // archive is where the seconds are (0.7 s against 12 s on a 39 MB file)
    expect(rendered.ariaNow).toBe('5')
    expect(rendered.barWidth).toBe('5%')
  })

  it('shows the write phase filling the bar', async () => {
    const rendered = await renderCard({ phase: 'write', fraction: 0.5 })
    // the write phase starts where the read ended (5%) and ends at 100%
    expect(rendered.ariaNow).toBe('53')
    expect(rendered.percent).toBe('53%')
    expect(rendered.barWidth).toBe('53%')
  })

  it('reaches 100% when the import is done', async () => {
    const rendered = await renderCard({ phase: 'write', fraction: 1 })
    expect(rendered.percent).toBe('100%')
    expect(rendered.ariaNow).toBe('100')
  })

  it('offers a cancel button that reaches the caller', async () => {
    const rendered = await renderCard({ phase: 'write', fraction: 0.2 })
    expect(rendered.cancelText).toBe('Cancel')
    const button = container?.querySelector<HTMLButtonElement>('.workbook-opening-cancel')
    expect(button).toBeTruthy()
    await act(async () => {
      button?.click()
    })
    expect(rendered.clicks()).toBe(1)
  })
})
