import { describe, expect, it, vi } from 'vitest'

import { runHeadlessPdfExport, waitForDeckSettled } from '../src/renderer/headless-export'
import type { DeckReadiness } from '../src/renderer/headless-export'

/**
 * Headless export readiness (src/renderer/headless-export.ts): rasterizing
 * before the deck's pictures have decoded would bake blank frames into the
 * PDF, so the hidden window waits — but never forever.
 */

const fast = { pollMs: 1, timeoutMs: 300 }

/** default readiness: a real deck is open and the open did not fail */
const deck = (over: Partial<DeckReadiness> = {}): DeckReadiness => ({
  slideCount: 1,
  pendingImages: 0,
  failed: false,
  ...over,
})

describe('waitForDeckSettled', () => {
  it('waits for pages and for every image to decode', async () => {
    let slideCount = 0
    let pendingImages = -1
    setTimeout(() => (slideCount = 4), 3)
    setTimeout(() => (pendingImages = 2), 5)
    setTimeout(() => (pendingImages = 0), 12)
    await waitForDeckSettled(() => deck({ slideCount, pendingImages }), fast)
    expect(slideCount).toBe(4)
    expect(pendingImages).toBe(0)
  })

  it('does not mistake "loader not created yet" for "no images pending"', async () => {
    // slideCount is already positive but the image effect has not run (-1)
    await expect(
      waitForDeckSettled(() => deck({ slideCount: 3, pendingImages: -1 }), fast),
    ).rejects.toThrow('images never finished decoding')
  })

  it('names the deck as the missing piece when nothing opened', async () => {
    await expect(waitForDeckSettled(() => deck({ slideCount: 0 }), fast)).rejects.toThrow(
      'did not open',
    )
  })

  it('fails immediately on the blank fallback instead of waiting out the budget', async () => {
    const t0 = Date.now()
    await expect(
      waitForDeckSettled(() => deck({ slideCount: 0, failed: true }), {
        pollMs: 1,
        timeoutMs: 5000,
      }),
    ).rejects.toThrow('did not open')
    expect(Date.now() - t0).toBeLessThan(1000)
  })

  it('waits for the private Office fonts to finish registering', async () => {
    window.__genofficeDocFontsSynced = false
    setTimeout(() => (window.__genofficeDocFontsSynced = true), 10)
    const t0 = Date.now()
    await waitForDeckSettled(deck, fast)
    expect(Date.now() - t0).toBeGreaterThanOrEqual(8)
    delete window.__genofficeDocFontsSynced
  })
})

describe('runHeadlessPdfExport', () => {
  it('runs the export once the deck has settled', async () => {
    const exportPdf = vi.fn(() => Promise.resolve(true))
    await expect(
      runHeadlessPdfExport('/o.pdf', () => deck({ slideCount: 2 }), exportPdf, fast),
    ).resolves.toEqual({ ok: true })
    expect(exportPdf).toHaveBeenCalledWith('/o.pdf')
  })

  it('reports the stall instead of exporting an empty deck', async () => {
    const exportPdf = vi.fn(() => Promise.resolve(true))
    await expect(
      runHeadlessPdfExport('/o.pdf', () => deck({ slideCount: 0 }), exportPdf, fast),
    ).resolves.toMatchObject({ ok: false, error: expect.stringContaining('did not open') })
    expect(exportPdf).not.toHaveBeenCalled()
  })
})
