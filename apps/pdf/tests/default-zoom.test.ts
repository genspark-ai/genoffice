import { describe, expect, it } from 'vitest'
import {
  LAST_ZOOM_KEY,
  loadLastZoom,
  resolveOpenZoom,
  sanitizePdfDefaultZoom,
  saveLastZoom,
} from '../src/renderer/view-state'
import type { PdfViewState } from '../src/renderer/view-state'

function memStorage(initial?: Record<string, string>) {
  const map = new Map(Object.entries(initial ?? {}))
  return {
    getItem: (k: string) => map.get(k) ?? null,
    setItem: (k: string, v: string) => void map.set(k, v),
    dump: () => map.get(LAST_ZOOM_KEY) ?? '',
  }
}

const saved = (over: Partial<PdfViewState> = {}): PdfViewState => ({
  page: 5,
  frac: 0.4,
  scale: 1.5,
  fitMode: null,
  ...over,
})

describe('sanitizePdfDefaultZoom', () => {
  it('keeps the three valid values, defaults everything else to width', () => {
    expect(sanitizePdfDefaultZoom('width')).toBe('width')
    expect(sanitizePdfDefaultZoom('page')).toBe('page')
    expect(sanitizePdfDefaultZoom('last')).toBe('last')
    expect(sanitizePdfDefaultZoom(undefined)).toBe('width')
    expect(sanitizePdfDefaultZoom('fit-page')).toBe('width')
    expect(sanitizePdfDefaultZoom('')).toBe('width')
  })
})

describe('resolveOpenZoom', () => {
  it('fit width always refits, ignoring any saved zoom', () => {
    expect(
      resolveOpenZoom({
        setting: 'width',
        savedView: saved(),
        lastZoom: { scale: 2, fitMode: null },
      }),
    ).toEqual({ fitMode: 'width', scale: null })
    expect(resolveOpenZoom({ setting: 'width', savedView: null, lastZoom: null })).toEqual({
      fitMode: 'width',
      scale: null,
    })
  })

  it('fit page always refits to the page, ignoring any saved zoom', () => {
    expect(
      resolveOpenZoom({
        setting: 'page',
        savedView: saved(),
        lastZoom: { scale: 3, fitMode: null },
      }),
    ).toEqual({ fitMode: 'page', scale: null })
  })

  it('last used restores the per-file zoom first', () => {
    expect(
      resolveOpenZoom({
        setting: 'last',
        savedView: saved(),
        lastZoom: { scale: 2, fitMode: null },
      }),
    ).toEqual({
      fitMode: null,
      scale: 1.5,
    })
    expect(
      resolveOpenZoom({
        setting: 'last',
        savedView: saved({ fitMode: 'page', scale: 1.23 }),
        lastZoom: { scale: 2, fitMode: null },
      }),
    ).toEqual({ fitMode: 'page', scale: null })
  })

  it('last used falls back to the global zoom for a never-opened file', () => {
    expect(
      resolveOpenZoom({ setting: 'last', savedView: null, lastZoom: { scale: 2, fitMode: null } }),
    ).toEqual({ fitMode: null, scale: 2 })
    expect(
      resolveOpenZoom({
        setting: 'last',
        savedView: null,
        lastZoom: { scale: 1.5, fitMode: 'width' },
      }),
    ).toEqual({ fitMode: 'width', scale: null })
  })

  it('last used with no history keeps the current behavior (fit width)', () => {
    expect(resolveOpenZoom({ setting: 'last', savedView: null, lastZoom: null })).toEqual({
      fitMode: 'width',
      scale: null,
    })
  })
})

describe('global last zoom storage', () => {
  it('round-trips the last used zoom', () => {
    const storage = memStorage()
    saveLastZoom({ scale: 1.75, fitMode: null }, storage)
    expect(loadLastZoom(storage)).toEqual({ scale: 1.75, fitMode: null })
  })

  it('returns null for missing or malformed entries', () => {
    expect(loadLastZoom(memStorage())).toBeNull()
    expect(loadLastZoom(memStorage({ [LAST_ZOOM_KEY]: '{bad' }))).toBeNull()
    expect(
      loadLastZoom(memStorage({ [LAST_ZOOM_KEY]: JSON.stringify({ scale: 'x', fitMode: null }) })),
    ).toBeNull()
    expect(
      loadLastZoom(memStorage({ [LAST_ZOOM_KEY]: JSON.stringify({ scale: 1, fitMode: 'zoom' }) })),
    ).toBeNull()
  })

  it('ignores setItem failures', () => {
    const storage = {
      getItem: () => null,
      setItem: () => {
        throw new Error('quota')
      },
    }
    expect(() => saveLastZoom({ scale: 1, fitMode: null }, storage)).not.toThrow()
  })
})
