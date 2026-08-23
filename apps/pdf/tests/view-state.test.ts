import { describe, expect, it } from 'vitest'
import {
  MAX_VIEW_ENTRIES,
  VIEW_STATE_KEY,
  captureViewState,
  loadViewState,
  saveViewState,
} from '../src/renderer/view-state'
import type { PdfViewState } from '../src/renderer/view-state'

function memStorage(initial?: Record<string, string>) {
  const map = new Map(Object.entries(initial ?? {}))
  return {
    getItem: (k: string) => map.get(k) ?? null,
    setItem: (k: string, v: string) => void map.set(k, v),
    dump: () => map.get(VIEW_STATE_KEY) ?? '',
  }
}

const state = (over: Partial<PdfViewState> = {}): PdfViewState => ({
  page: 5,
  frac: 0.4,
  scale: 1.25,
  fitMode: null,
  ...over,
})

describe('saveViewState / loadViewState', () => {
  it('round-trips a saved position', () => {
    const storage = memStorage()
    saveViewState('/a.pdf', state(), storage)
    expect(loadViewState('/a.pdf', storage)).toEqual(state())
  })

  it('returns null for unknown paths and empty paths', () => {
    const storage = memStorage()
    saveViewState('/a.pdf', state(), storage)
    expect(loadViewState('/other.pdf', storage)).toBeNull()
    expect(loadViewState('', storage)).toBeNull()
  })

  it('keeps entries for other files when saving', () => {
    const storage = memStorage()
    saveViewState('/a.pdf', state({ page: 2 }), storage)
    saveViewState('/b.pdf', state({ page: 9 }), storage)
    expect(loadViewState('/a.pdf', storage)?.page).toBe(2)
    expect(loadViewState('/b.pdf', storage)?.page).toBe(9)
  })

  it('survives corrupted store JSON', () => {
    const storage = memStorage({ [VIEW_STATE_KEY]: '{not json' })
    expect(loadViewState('/a.pdf', storage)).toBeNull()
    saveViewState('/a.pdf', state(), storage)
    expect(loadViewState('/a.pdf', storage)).toEqual(state())
  })

  it('rejects malformed entries', () => {
    const bad = {
      '/no-page.pdf': { frac: 0, scale: 1, fitMode: null, at: 1 },
      '/zero-page.pdf': { page: 0, frac: 0, scale: 1, fitMode: null, at: 1 },
      '/nan-scale.pdf': { page: 1, frac: 0, scale: 'x', fitMode: null, at: 1 },
      '/neg-scale.pdf': { page: 1, frac: 0, scale: -1, fitMode: null, at: 1 },
      '/bad-fit.pdf': { page: 1, frac: 0, scale: 1, fitMode: 'zoom', at: 1 },
      '/not-object.pdf': 42,
    }
    const storage = memStorage({ [VIEW_STATE_KEY]: JSON.stringify(bad) })
    for (const path of Object.keys(bad)) expect(loadViewState(path, storage)).toBeNull()
  })

  it('clamps a stored frac into [0, 1] and defaults a missing frac to 0', () => {
    const storage = memStorage({
      [VIEW_STATE_KEY]: JSON.stringify({
        '/over.pdf': { page: 1, frac: 7, scale: 1, fitMode: 'width', at: 1 },
        '/missing.pdf': { page: 1, scale: 1, fitMode: 'width', at: 1 },
      }),
    })
    expect(loadViewState('/over.pdf', storage)?.frac).toBe(1)
    expect(loadViewState('/missing.pdf', storage)?.frac).toBe(0)
  })

  it('prunes the least recently saved entries beyond the cap', () => {
    const storage = memStorage()
    for (let i = 0; i < MAX_VIEW_ENTRIES; i++)
      saveViewState(`/f${i}.pdf`, state({ page: i + 1 }), storage, 1000 + i)
    // Re-saving an old entry refreshes it, so it must survive the prune below
    saveViewState('/f0.pdf', state({ page: 1 }), storage, 5000)
    saveViewState('/new.pdf', state(), storage, 6000)
    expect(loadViewState('/new.pdf', storage)).not.toBeNull()
    expect(loadViewState('/f0.pdf', storage)).not.toBeNull()
    // f1 is now the oldest and the only one over the cap
    expect(loadViewState('/f1.pdf', storage)).toBeNull()
    expect(loadViewState('/f2.pdf', storage)).not.toBeNull()
  })

  it('ignores setItem failures (quota exceeded)', () => {
    const storage = {
      getItem: () => null,
      setItem: () => {
        throw new Error('quota')
      },
    }
    expect(() => saveViewState('/a.pdf', state(), storage)).not.toThrow()
  })
})

describe('captureViewState', () => {
  // Layout: gap 10, rows of 100px → row tops at 10, 120, 230, ...
  const geometry = {
    rowHeights: [100, 100, 100],
    rowPages: [1, 2, 3],
    gap: 10,
    scale: 1.5,
    fitMode: 'width' as const,
  }

  it('captures the row at the viewport top with its within-row fraction', () => {
    const view = captureViewState({ ...geometry, scrollTop: 170 })
    expect(view).toEqual({ page: 2, frac: 0.5, scale: 1.5, fitMode: 'width' })
  })

  it('captures the document top as page 1 frac 0', () => {
    expect(captureViewState({ ...geometry, scrollTop: 0 })).toMatchObject({ page: 1, frac: 0 })
  })

  it('clamps a scroll inside the inter-row gap to frac 1', () => {
    // 115 is past row 0's content (10..110) but before row 1's top (120)
    expect(captureViewState({ ...geometry, scrollTop: 115 })).toMatchObject({ page: 1, frac: 1 })
  })

  it('sticks to the last row when overscrolled', () => {
    expect(captureViewState({ ...geometry, scrollTop: 9999 })).toMatchObject({ page: 3, frac: 1 })
  })

  it('handles an empty document without dividing by zero', () => {
    const view = captureViewState({
      rowHeights: [],
      rowPages: [],
      gap: 10,
      scale: 1,
      fitMode: null,
      scrollTop: 50,
    })
    expect(view).toMatchObject({ page: 1, frac: 0 })
  })

  it('carries spread-mode row pages through', () => {
    const view = captureViewState({
      rowHeights: [100, 100],
      rowPages: [1, 3],
      gap: 10,
      scale: 1,
      fitMode: null,
      scrollTop: 120,
    })
    expect(view.page).toBe(3)
  })
})
